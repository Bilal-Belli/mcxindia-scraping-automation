
"""Merge MCX CSV files into a sorted master CSV and synchronize Access."""

from pathlib import Path

import pandas as pd
import pypyodbc

BASE_DIR = Path(__file__).resolve().parent
DB_FILE = BASE_DIR / "database.mdb"
INPUT_DIR = BASE_DIR / "filtredData"
MASTER_FILE = BASE_DIR / "combined_sorted_mcx.csv"

COLS = [
    "Expiry_dt",
    "Open",
    "High",
    "Low",
    "Close",
    "Val_in",
    "Open_int",
    "Daily",
]

MASTER_COLS = ["Symbol"] + COLS

SUFFIXES = {
    1: "_ONE",
    2: "_TWO",
    3: "_THREE",
}

SOURCE_COLUMNS = {
    "Symbol",
    "Date",
    "Expiry Date",
    "Open",
    "High",
    "Low",
    "Close",
    "Value(Lacs)",
    "Open Interest(Lots)",
}


def read_formatted_file(path: Path) -> pd.DataFrame:
    """Convert a filtered MCX CSV into the master CSV schema."""

    df = pd.read_csv(
        path,
        dtype=str,
        keep_default_na=False,
        encoding="utf-8-sig",
    )

    df.columns = df.columns.str.strip()

    missing = SOURCE_COLUMNS - set(df.columns)

    if missing:
        raise ValueError(
            f"{path.name}: missing columns {sorted(missing)}"
        )

    daily = pd.to_datetime(
        df["Date"].str.strip(),
        format="%d %b %Y",
        errors="coerce",
    )

    expiry = pd.to_datetime(
        df["Expiry Date"].str.strip(),
        format="%d%b%Y",
        errors="coerce",
    )

    result = pd.DataFrame({
        "Symbol": df["Symbol"].str.strip(),
        "Expiry_dt": expiry.dt.strftime("%d-%m-%Y"),
        "Open": df["Open"],
        "High": df["High"],
        "Low": df["Low"],
        "Close": df["Close"],
        "Val_in": df["Value(Lacs)"],
        "Open_int": df["Open Interest(Lots)"],
        "Daily": daily.dt.strftime("%d-%m-%Y"),
    })

    valid = (
        daily.notna()
        & expiry.notna()
        & result["Symbol"].ne("")
    )

    return result.loc[valid, MASTER_COLS].copy()


def merge_master(
    files: list[Path] | None = None,
) -> pd.DataFrame:
    """
    Load the master CSV if present.
    Otherwise, initialize it from all filtered MCX CSV files.
    Merge incoming files, deduplicate, and sort chronologically.
    """

    if MASTER_FILE.exists():
        print(f"-> Loading master: {MASTER_FILE.name}")

        master = pd.read_csv(
            MASTER_FILE,
            dtype=str,
            keep_default_na=False,
            encoding="utf-8-sig",
        )

        missing = set(MASTER_COLS) - set(master.columns)

        if missing:
            raise ValueError(
                f"Master CSV missing columns: {sorted(missing)}"
            )

        master = master[MASTER_COLS].copy()

        files_to_merge = sorted(
            Path(file)
            for file in (files or [])
            if Path(file).exists()
        )

    else:
        print(
            "-> Master CSV not found. "
            "Building it from all filtered MCX CSV files."
        )

        files_to_merge = sorted(INPUT_DIR.glob("MCX*.csv"))
        master = pd.DataFrame(columns=MASTER_COLS)

        if not files_to_merge:
            raise FileNotFoundError(
                f"No filtered CSV files found in {INPUT_DIR}"
            )

    frames = [master]
    successful_files = 0

    for path in files_to_merge:
        print(f"-> Reading {path.name}")

        try:
            frame = read_formatted_file(path)

            if not frame.empty:
                frames.append(frame)
                successful_files += 1
            else:
                print("   [Warning] No valid records found.")

        except Exception as exc:
            print(f"   [Warning] Skipping {path.name}: {exc}")

    if not MASTER_FILE.exists() and successful_files == 0:
        raise ValueError(
            "Could not initialize the master CSV: "
            "no valid records were found."
        )

    if MASTER_FILE.exists() and not files_to_merge:
        print("-> No new CSV files to merge.")

    df = pd.concat(frames, ignore_index=True)

    if df.empty:
        raise ValueError("No valid records available for the master CSV.")

    for column in MASTER_COLS:
        df[column] = (
            df[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    daily = pd.to_datetime(
        df["Daily"],
        format="%d-%m-%Y",
        errors="coerce",
    )

    expiry = pd.to_datetime(
        df["Expiry_dt"],
        format="%d-%m-%Y",
        errors="coerce",
    )

    valid = (
        daily.notna()
        & expiry.notna()
        & df["Symbol"].ne("")
    )

    df = df.loc[valid].copy()
    df["_daily_sort"] = daily.loc[valid]
    df["_expiry_sort"] = expiry.loc[valid]

    # One record per symbol, trading day, and expiry.
    # New incoming records take precedence over existing records.
    df = df.drop_duplicates(
        subset=["Symbol", "Daily", "Expiry_dt"],
        keep="last",
    )

    df = df.sort_values(
        ["_daily_sort", "_expiry_sort", "Symbol"],
        kind="stable",
    )

    master = df[MASTER_COLS].reset_index(drop=True)

    master.to_csv(
        MASTER_FILE,
        index=False,
        encoding="utf-8-sig",
    )

    print(f"-> Master contains {len(master):,} records.")
    print(f"-> Saved: {MASTER_FILE.name}")

    return master


def table_exists(cursor, table_name: str) -> bool:
    """Check whether an Access table exists."""

    cursor.tables(table=table_name)
    return bool(cursor.fetchall())


def prepare_tables(
    master: pd.DataFrame,
) -> dict[str, pd.DataFrame]:
    """
    Assign each row to ONE, TWO, or THREE according to
    its expiry rank for that particular trading day and symbol.
    """

    df = master.copy()

    df["_daily_sort"] = pd.to_datetime(
        df["Daily"],
        format="%d-%m-%Y",
        errors="coerce",
    )

    df["_expiry_sort"] = pd.to_datetime(
        df["Expiry_dt"],
        format="%d-%m-%Y",
        errors="coerce",
    )

    df = df[
        df["_daily_sort"].notna()
        & df["_expiry_sort"].notna()
    ].copy()

    if df.empty:
        return {}

    df["_rank"] = (
        df.groupby(
            ["Daily", "Symbol"],
            sort=False,
        )["_expiry_sort"]
        .rank(method="dense")
    )

    df = df[df["_rank"].isin([1, 2, 3])].copy()

    if df.empty:
        return {}

    df["_rank"] = df["_rank"].astype(int)

    df["_table"] = df["Symbol"] + df["_rank"].map(SUFFIXES)

    invalid = ~df["_table"].str.fullmatch(
        r"[A-Za-z0-9_]+",
        na=False,
    )

    if invalid.any():
        bad_names = df.loc[invalid, "_table"].unique()[:10]

        raise ValueError(
            f"Invalid Access table names: {list(bad_names)}"
        )

    return {
        name: group[COLS].copy()
        for name, group in df.groupby("_table", sort=False)
    }


def synchronize_mdb(master: pd.DataFrame) -> None:
    """Insert missing records from the master CSV into Access."""

    if master.empty:
        print("No valid records to synchronize.")
        return

    tables = prepare_tables(master)

    if not tables:
        print("No eligible expiry tables found.")
        return

    if not DB_FILE.exists():
        print(f"-> Creating Access database: {DB_FILE.name}")
        pypyodbc.win_create_mdb(str(DB_FILE))

    conn = None
    cursor = None
    inserted_total = 0
    skipped_total = 0

    try:
        conn = pypyodbc.connect(
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={DB_FILE};"
        )

        cursor = conn.cursor()

        for table_name, table_df in tables.items():
            print(f"-> Synchronizing {table_name}")

            if not table_exists(cursor, table_name):
                definitions = ", ".join(
                    f"[{column}] TEXT"
                    for column in COLS
                )

                cursor.execute(
                    f"CREATE TABLE [{table_name}] ({definitions})"
                )

                # Commit table creation separately for Access/ODBC.
                conn.commit()
                existing_keys = set()

            else:
                cursor.execute(
                    f"SELECT [Daily], [Expiry_dt] "
                    f"FROM [{table_name}]"
                )

                existing_keys = {
                    (str(row[0]).strip(), str(row[1]).strip())
                    for row in cursor.fetchall()
                    if row[0] is not None and row[1] is not None
                }

            # One record per trading day and expiry in each table.
            candidates = table_df.drop_duplicates(
                subset=["Daily", "Expiry_dt"],
                keep="last",
            )

            missing_rows = []

            for row in candidates[COLS].itertuples(
                index=False,
                name=None,
            ):
                daily_value = str(row[COLS.index("Daily")]).strip()
                expiry_value = str(
                    row[COLS.index("Expiry_dt")]
                ).strip()

                key = (daily_value, expiry_value)

                if key in existing_keys:
                    continue

                missing_rows.append(
                    tuple(
                        "" if pd.isna(value) else str(value)
                        for value in row
                    )
                )

                existing_keys.add(key)

            if not missing_rows:
                skipped_total += len(candidates)
                continue

            columns_sql = ", ".join(
                f"[{column}]" for column in COLS
            )

            placeholders = ", ".join("?" for _ in COLS)

            insert_sql = (
                f"INSERT INTO [{table_name}] "
                f"({columns_sql}) VALUES ({placeholders})"
            )

            cursor.executemany(insert_sql, missing_rows)

            inserted_total += len(missing_rows)
            skipped_total += len(candidates) - len(missing_rows)

        # Commit all insertions together.
        conn.commit()

        print("\n" + "=" * 55)
        print("MASTER CSV -> ACCESS SYNCHRONIZATION COMPLETE")
        print("=" * 55)
        print(f"Tables checked: {len(tables):,}")
        print(f"Rows inserted:  {inserted_total:,}")
        print(f"Rows skipped:   {skipped_total:,}")
        print(f"Master CSV:     {MASTER_FILE.name}")
        print(f"Access DB:      {DB_FILE.name}")

    except Exception:
        if conn is not None:
            try:
                conn.rollback()
            except Exception:
                pass

        raise

    finally:
        if cursor is not None:
            try:
                cursor.close()
            except Exception:
                pass

        if conn is not None:
            try:
                conn.close()
            except Exception as exc:
                print(f"[Warning] Connection close failed: {exc}")


def process_to_mdb(
    files: list[Path] | None = None,
) -> None:
    """
    Merge new files into the master CSV, then synchronize Access.

    If files is None, scan all filtered CSV files. This is useful
    for the first run or recovery after an interrupted pipeline.
    """

    if files is None:
        files = sorted(INPUT_DIR.glob("MCX*.csv"))
    else:
        files = sorted(
            Path(file)
            for file in files
            if Path(file).exists()
        )

    if not files and not MASTER_FILE.exists():
        raise FileNotFoundError(
            "Neither the master CSV nor filtered CSV files exist."
        )

    master = merge_master(files)
    synchronize_mdb(master)


if __name__ == "__main__":
    try:
        process_to_mdb()
    except Exception as exc:
        print(f"\n[ERROR] {type(exc).__name__}: {exc}")
