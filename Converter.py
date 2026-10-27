"""Converter: new CSV files -> sorted Master CSV -> Access MDB.

Flow:
  1. Combine the new filtered CSV files into the master CSV (dedupe, sort by Daily).
  2. Load the master into Access with real DATETIME columns (Daily, Expiry_dt),
     so rows are sorted by date (primary key = Daily + Expiry_dt).
  3. Show those dates as dd-mm-yyyy in Access.
"""

from datetime import datetime
from pathlib import Path

import pandas as pd
import pypyodbc

try:
    import win32com.client
except ImportError:  # only needed for the dd-mm-yyyy display format
    win32com = None

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


# --------------------------------------------------------------------------
# Step 1: combine new files into the master CSV
# --------------------------------------------------------------------------

def read_formatted_file(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    df.columns = df.columns.str.strip()

    missing = SOURCE_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(f"{path.name}: missing columns {sorted(missing)}")

    daily = pd.to_datetime(df["Date"].str.strip(), format="%d %b %Y", errors="coerce")
    expiry = pd.to_datetime(df["Expiry Date"].str.strip(), format="%d%b%Y", errors="coerce")

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

    valid = daily.notna() & expiry.notna() & result["Symbol"].ne("")
    return result.loc[valid, MASTER_COLS].copy()


def merge_master(files: list[Path]) -> pd.DataFrame:
    """Merge the given files into the master CSV and save it sorted by Daily."""

    if MASTER_FILE.exists():
        print(f"-> Loading master: {MASTER_FILE.name}")
        master = pd.read_csv(MASTER_FILE, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        missing = set(MASTER_COLS) - set(master.columns)
        if missing:
            raise ValueError(f"Master CSV missing columns: {sorted(missing)}")
        master = master[MASTER_COLS].copy()
    else:
        print("-> Master CSV not found. Building it from the filtered CSV files.")
        master = pd.DataFrame(columns=MASTER_COLS)
        if not files:
            raise FileNotFoundError(f"No filtered CSV files found in {INPUT_DIR}")

    frames = [master]
    for path in files:
        print(f"-> Reading {path.name}")
        try:
            frame = read_formatted_file(path)
            if not frame.empty:
                frames.append(frame)
            else:
                print("   [Warning] No valid records found.")
        except Exception as exc:
            print(f"   [Warning] Skipping {path.name}: {exc}")

    df = pd.concat(frames, ignore_index=True)
    if df.empty:
        raise ValueError("No valid records available.")

    for column in MASTER_COLS:
        df[column] = df[column].fillna("").astype(str).str.strip()

    non_empty_mask = (df[MASTER_COLS] != "").all(axis=1)
    df = df.loc[non_empty_mask].copy()

    daily = pd.to_datetime(df["Daily"], format="%d-%m-%Y", errors="coerce")
    expiry = pd.to_datetime(df["Expiry_dt"], format="%d-%m-%Y", errors="coerce")

    valid = daily.notna() & expiry.notna() & df["Symbol"].ne("")
    df = df.loc[valid].copy()
    df["_daily_sort"] = daily.loc[valid]
    df["_expiry_sort"] = expiry.loc[valid]

    # New incoming records take precedence over existing ones.
    df = df.drop_duplicates(subset=["Symbol", "Daily", "Expiry_dt"], keep="last")
    df = df.sort_values(["_daily_sort", "_expiry_sort", "Symbol"], kind="stable")

    master_sorted = df[MASTER_COLS].reset_index(drop=True)

    print(f"-> Saving sorted Master CSV: {MASTER_FILE.name} ({len(master_sorted):,} records)")
    master_sorted.to_csv(MASTER_FILE, index=False, encoding="utf-8-sig")

    return master_sorted


# --------------------------------------------------------------------------
# Step 2 + 3: Access with real dates, shown as dd-mm-yyyy
# --------------------------------------------------------------------------

def to_date(value):
    """Return a datetime.date from a datetime, 'YYYY-MM-DD', 'DD-MM-YYYY' or an Access string.

    Returns None if the value is empty or cannot be parsed.
    """
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if hasattr(value, "date") and callable(value.date):
        return value.date()
    if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
        return value  # already a date

    s = str(value).strip().split()[0] if str(value).strip() else ""
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def table_exists(cursor, table_name: str) -> bool:
    cursor.tables(table=table_name)
    return bool(cursor.fetchall())


def set_access_date_formats(mdb_path: Path, table_names: list[str]) -> None:
    """Uses Access DAO ONCE to set the display format of date fields to dd-mm-yyyy for all newly created tables."""
    if not table_names:
        return
    if win32com is None:
        print("   [Note] pywin32 not installed: dd-mm-yyyy display format not applied.")
        return

    print(f"-> Applying 'dd-mm-yyyy' display format to {len(table_names)} table(s) via DAO...")
    try:
        access_app = win32com.client.Dispatch("Access.Application")
        access_app.OpenCurrentDatabase(str(mdb_path))
        db = access_app.CurrentDb()

        for table_name in table_names:
            try:
                table_def = db.TableDefs(table_name)
                for col_name in ["Expiry_dt", "Daily"]:
                    field = table_def.Fields(col_name)
                    try:
                        field.Properties("Format").Value = "dd-mm-yyyy"
                    except Exception:
                        prop = field.CreateProperty("Format", 10, "dd-mm-yyyy")
                        field.Properties.Append(prop)
            except Exception as e:
                print(f"   [Note] Could not apply format to {table_name}: {e}")

        access_app.Quit()
        print("-> Date formats applied successfully!")
    except Exception as e:
        print(f"   [Note] Could not open Access via DAO for formatting: {e}")


def synchronize_mdb(master: pd.DataFrame) -> None:
    print("-> Synchronizing into Access Database with True Date/Time Types...")

    df = master.copy()

    dt_daily = pd.to_datetime(df["Daily"], format="%d-%m-%Y", errors="coerce")
    dt_expiry = pd.to_datetime(df["Expiry_dt"], format="%d-%m-%Y", errors="coerce")

    df["Expiry_ISO"] = dt_expiry.dt.strftime("%Y-%m-%d")
    df["Daily_ISO"] = dt_daily.dt.strftime("%Y-%m-%d")

    df["_expiry_sort"] = dt_expiry
    df["_rank"] = df.groupby(["Daily", "Symbol"], sort=False)["_expiry_sort"].rank(method="dense")
    df = df[df["_rank"].isin([1, 2, 3])].copy()
    df["_rank"] = df["_rank"].astype(int)
    # Access table names are case-insensitive, so normalise to upper case.
    df["_table"] = df["Symbol"].str.upper() + df["_rank"].map(SUFFIXES)

    tables = {name: group.copy() for name, group in df.groupby("_table", sort=False)}

    if not DB_FILE.exists():
        pypyodbc.win_create_mdb(str(DB_FILE))

    conn = pypyodbc.connect(r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};" f"DBQ={DB_FILE};")
    cursor = conn.cursor()

    inserted_total = 0
    skipped_total = 0
    newly_created_tables = []

    db_cols = ["Expiry_ISO", "Open", "High", "Low", "Close", "Val_in", "Open_int", "Daily_ISO"]

    for table_name, table_df in tables.items():
        existing_keys = set()

        if not table_exists(cursor, table_name):
            definitions = []
            for column in COLS:
                if column in ("Expiry_dt", "Daily"):
                    definitions.append(f"[{column}] DATETIME")
                else:
                    definitions.append(f"[{column}] TEXT")

            primary_key_def = f", CONSTRAINT [PK_{table_name}] PRIMARY KEY ([Daily], [Expiry_dt])"
            cursor.execute(f"CREATE TABLE [{table_name}] ({', '.join(definitions)}{primary_key_def})")
            conn.commit()

            newly_created_tables.append(table_name)
        else:
            cursor.execute(f"SELECT [Daily], [Expiry_dt] FROM [{table_name}]")
            for row in cursor.fetchall():
                d_day, d_exp = to_date(row[0]), to_date(row[1])
                if d_day and d_exp:
                    existing_keys.add((d_day, d_exp))

        candidates = table_df.drop_duplicates(subset=["Daily_ISO", "Expiry_ISO"], keep="last")
        missing_rows = []

        for exp_s, o, h, l, c, v, oi, day_s in candidates[db_cols].itertuples(index=False, name=None):
            exp_d, day_d = to_date(exp_s), to_date(day_s)
            if exp_d is None or day_d is None:
                continue

            key = (day_d, exp_d)
            if key in existing_keys:
                continue
            existing_keys.add(key)

            missing_rows.append((
                exp_d,
                "" if pd.isna(o) else str(o),
                "" if pd.isna(h) else str(h),
                "" if pd.isna(l) else str(l),
                "" if pd.isna(c) else str(c),
                "" if pd.isna(v) else str(v),
                "" if pd.isna(oi) else str(oi),
                day_d,
            ))

        if not missing_rows:
            skipped_total += len(candidates)
            continue

        columns_sql = ", ".join(f"[{c}]" for c in COLS)

        try:
            for exp_d, o, h, l, c, v, oi, day_d in missing_rows:
                # Access/Jet date literals are always read as US mm/dd/yyyy, whatever the PC locale.
                insert_sql = (
                    f"INSERT INTO [{table_name}] ({columns_sql}) VALUES ("
                    f"#{exp_d.month:02d}/{exp_d.day:02d}/{exp_d.year}#, ?, ?, ?, ?, ?, ?, "
                    f"#{day_d.month:02d}/{day_d.day:02d}/{day_d.year}#)"
                )
                cursor.execute(insert_sql, (o, h, l, c, v, oi))
            conn.commit()
        except Exception as exc:
            conn.rollback()
            print(f"   [Error] Table {table_name}: {exc}")
            raise

        inserted_total += len(missing_rows)
        skipped_total += len(candidates) - len(missing_rows)

    cursor.close()
    conn.close()

    if newly_created_tables:
        set_access_date_formats(DB_FILE, newly_created_tables)

    print(f"\n[SUCCESS] MDB Synchronized! Inserted: {inserted_total:,} rows, Skipped: {skipped_total:,} rows.")


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------

def process_to_mdb(files: list[Path] | None = None) -> None:
    """Merge new files into the master CSV, then sync Access (real dates, dd-mm-yyyy).

    If files is None, all filtered MCX CSV files in INPUT_DIR are scanned.
    Already-known rows are deduplicated, so re-reading old files is harmless.
    """
    print("-> Starting Master Pipeline...")

    if files is None:
        files = sorted(INPUT_DIR.glob("MCX*.csv"))
    else:
        files = sorted(Path(f) for f in files if Path(f).exists())

    if not files and not MASTER_FILE.exists():
        raise FileNotFoundError("Neither the master CSV nor filtered CSV files exist.")

    if not files:
        print("-> No new CSV files to merge.")

    master = merge_master(files)
    synchronize_mdb(master)


run_pipeline = process_to_mdb  # backward-compatible name


if __name__ == "__main__":
    try:
        process_to_mdb()
    except Exception as exc:
        print(f"\n[ERROR] {type(exc).__name__}: {exc}")