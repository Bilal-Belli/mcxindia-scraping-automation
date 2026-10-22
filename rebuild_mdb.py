
from pathlib import Path
import shutil
import pandas as pd
import pypyodbc

BASE_DIR = Path(__file__).resolve().parent
DB_FILE = BASE_DIR / "database.mdb"
INPUT_DIR = BASE_DIR / "filtredData"
COMBINED_CSV_FILE = BASE_DIR / "combined_sorted_mcx.csv"
TMP_DIR = BASE_DIR / "_tmp_import"

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

SUFFIXES = {1: "_ONE", 2: "_TWO", 3: "_THREE"}


def load_master() -> pd.DataFrame | None:
    """Load the combined CSV or build it from the filtered source files."""

    if COMBINED_CSV_FILE.exists():
        print(f"-> Loading {COMBINED_CSV_FILE.name}")

        df = pd.read_csv(
            COMBINED_CSV_FILE,
            dtype=str,
            keep_default_na=False,
        )

        required = {"Symbol", *COLS}
        missing = required - set(df.columns)

        if missing:
            raise ValueError(
                f"Missing columns in {COMBINED_CSV_FILE.name}: "
                f"{', '.join(sorted(missing))}"
            )

        df["Symbol"] = df["Symbol"].str.strip()
        df["_dt_expiry"] = pd.to_datetime(
            df["Expiry_dt"],
            format="%d-%m-%Y",
            errors="coerce",
        )

        return df

    files = sorted(INPUT_DIR.glob("MCX*.csv"))

    if not files:
        print(f"No filtered CSV files found in: {INPUT_DIR}")
        return None

    dfs = []

    for file in files:
        try:
            data = pd.read_csv(
                file,
                dtype=str,
                keep_default_na=False,
            )

            if not data.empty:
                dfs.append(data)

        except Exception as exc:
            print(f"  [Warning] Could not read {file.name}: {exc}")

    if not dfs:
        print("No valid data found.")
        return None

    df = pd.concat(dfs, ignore_index=True)

    required = {
        "Symbol",
        "Date",
        "Expiry Date",
        "Value(Lacs)",
        "Open Interest(Lots)",
        "Open",
        "High",
        "Low",
        "Close",
    }

    missing = required - set(df.columns)

    if missing:
        raise ValueError(
            "Missing source columns: " + ", ".join(sorted(missing))
        )

    df["Symbol"] = df["Symbol"].str.strip()

    daily = pd.to_datetime(
        df["Date"],
        format="%d %b %Y",
        errors="coerce",
    )

    expiry = pd.to_datetime(
        df["Expiry Date"],
        format="%d%b%Y",
        errors="coerce",
    )

    df["Daily"] = daily.dt.strftime("%d-%m-%Y")
    df["Expiry_dt"] = expiry.dt.strftime("%d-%m-%Y")

    df["_dt_daily"] = daily
    df["_dt_expiry"] = expiry

    df = df.rename(
        columns={
            "Value(Lacs)": "Val_in",
            "Open Interest(Lots)": "Open_int",
        }
    )

    df = df[["Symbol"] + COLS + ["_dt_daily", "_dt_expiry"]]

    df = df.sort_values(
        ["_dt_daily", "_dt_expiry"],
        kind="stable",
        na_position="last",
    )

    # Save the combined CSV without internal sorting columns.
    df[["Symbol"] + COLS].to_csv(
        COMBINED_CSV_FILE,
        index=False,
    )

    print(f"-> Created {COMBINED_CSV_FILE.name}")

    return df.drop(columns=["_dt_daily"])


def rebuild_mdb_safely() -> None:
    """Rebuild the Access database from the combined market data."""

    df = load_master()

    if df is None or df.empty:
        print("No data to import.")
        return

    # Rank distinct expiry dates for each trading day and symbol.
    df["_rank"] = (
        df.groupby(
            ["Daily", "Symbol"],
            sort=False,
        )["_dt_expiry"]
        .rank(method="dense")
    )

    # Keep only the first three distinct expiry dates.
    df = df[df["_rank"].isin([1, 2, 3])].copy()

    # Rows without a valid expiry date cannot be assigned a rank.
    df = df[df["_rank"].notna()].copy()

    if df.empty:
        print("No valid expiry data to import.")
        return

    df["_rank"] = df["_rank"].astype(int)

    # Validate table names before using them in SQL statements.
    df["_table"] = (
        df["Symbol"] + df["_rank"].map(SUFFIXES)
    )

    invalid = df[
        ~df["_table"].str.fullmatch(r"[A-Za-z0-9_]+", na=False)
    ]

    if not invalid.empty:
        raise ValueError(
            "Invalid Access table names found: "
            + ", ".join(invalid["_table"].unique()[:10])
        )

    # Prepare temporary CSV files and schema.ini.
    if TMP_DIR.exists():
        shutil.rmtree(TMP_DIR)

    TMP_DIR.mkdir(parents=True)

    tables = {}
    schema = []

    try:
        for i, (name, group) in enumerate(
            df.groupby("_table", sort=False)
        ):
            filename = f"t{i}.csv"
            csv_path = TMP_DIR / filename

            # Force every imported column to text.
            group[COLS].to_csv(
                csv_path,
                index=False,
                sep=";",
                encoding="cp1252",
            )

            tables[name] = filename

            schema.extend(
                [
                    f"[{filename}]",
                    "Format=Delimited(;)",
                    "ColNameHeader=True",
                    "DecimalSymbol=.",
                    "CharacterSet=ANSI",
                ]
            )

            schema.extend(
                f"Col{j + 1}={column} Text"
                for j, column in enumerate(COLS)
            )

            schema.append("")

        (TMP_DIR / "schema.ini").write_text(
            "\n".join(schema),
            encoding="ascii",
        )

        # Create a fresh Access database.
        if DB_FILE.exists():
            DB_FILE.unlink()

        print(f"-> Creating {DB_FILE.name}")
        pypyodbc.win_create_mdb(str(DB_FILE))

        connection_string = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={DB_FILE};"
        )

        conn = None
        cur = None

        try:
            conn = pypyodbc.connect(connection_string)
            cur = conn.cursor()

            # Do not enable autocommit.
            source = f"[Text;HDR=Yes;DATABASE={TMP_DIR}]"

            for name, filename in tables.items():
                print(f"-> Importing table: {name}")

                sql = (
                    f"SELECT * INTO [{name}] "
                    f"FROM {source}.[{filename}]"
                )

                cur.execute(sql)

            # Commit after all tables have been imported.
            conn.commit()

            print(f"Done: {len(tables)} tables built.")
            print(f"Database: {DB_FILE}")

        except Exception:
            if conn is not None:
                try:
                    conn.rollback()
                except Exception:
                    pass

            raise

        finally:
            if cur is not None:
                try:
                    cur.close()
                except Exception:
                    pass

            if conn is not None:
                try:
                    conn.close()
                except Exception as exc:
                    print(
                        f"[Warning] Connection close failed: {exc}"
                    )

    finally:
        # Remove temporary files whether the import succeeds or fails.
        shutil.rmtree(TMP_DIR, ignore_errors=True)


if __name__ == "__main__":
    try:
        rebuild_mdb_safely()
    except Exception as exc:
        print(f"\n[ERROR] {type(exc).__name__}: {exc}")

    input("\nPress Enter to close this window...")
