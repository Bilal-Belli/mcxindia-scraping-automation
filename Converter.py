"""Standalone script to process all filtered MCX CSV files into an Access database (.mdb) at maximum speed."""

from pathlib import Path
import pandas as pd
import pypyodbc

BASE_DIR = Path(__file__).resolve().parent
DB_FILE = BASE_DIR / 'database.mdb'
INPUT_DIR = BASE_DIR / 'filtredData'

def process_to_mdb(files: list[Path] = None) -> None:
    DB_FILE.parent.mkdir(parents=True, exist_ok=True)

    if not DB_FILE.exists():
        pypyodbc.win_create_mdb(str(DB_FILE))

    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={DB_FILE};'
    )
    conn = pypyodbc.connect(conn_str)
    cursor = conn.cursor()
    
    conn.autocommit = False

    def table_exists(cur, t_name):
        try:
            cur.tables(table=t_name)
            return len(cur.fetchall()) > 0
        except Exception:
            return False

    if files is None:
        files = sorted(INPUT_DIR.glob('MCX*.csv'))
    else:
        files = sorted([f for f in files if f.exists()])

    if not files:
        print(f"No filtered CSV files found to process into MDB.")
        conn.close()
        return

    print(f"Processing {len(files)} filtered file(s) into MDB at high speed.\n")
    suffixes = ['_ONE', '_TWO', '_THREE']

    for csv_file in files:
        print(f"Processing file: {csv_file.name}")
        try:
            df = pd.read_csv(csv_file, dtype=str, keep_default_na=False)
            if df.empty:
                continue

            df['Symbol'] = df['Symbol'].str.strip()

            dt_date = pd.to_datetime(df['Date'], format='%d %b %Y', errors='coerce')
            df['Daily'] = dt_date.dt.strftime('%d-%m-%Y')

            dt_expiry = pd.to_datetime(df['Expiry Date'], format='%d%b%Y', errors='coerce')
            df['Expiry_dt'] = dt_expiry.dt.strftime('%d-%m-%Y')
            df['_raw_expiry'] = dt_expiry

            df = df.rename(columns={'Value(Lacs)': 'Val_in', 'Open Interest(Lots)': 'Open_int'})

            keep_cols = ['Symbol', 'Expiry_dt', 'Open', 'High', 'Low', 'Close', 'Val_in', 'Open_int', 'Daily', '_raw_expiry']
            df = df[[c for c in keep_cols if c in df.columns]]

            for symbol, sym_df in df.groupby('Symbol'):
                unique_expiries = sorted(sym_df['_raw_expiry'].dropna().unique())

                for i, exp_date in enumerate(unique_expiries[:3]):
                    if i >= len(suffixes):
                        break
                    
                    table_name = f'{symbol}{suffixes[i]}'
                    sub_df = sym_df[sym_df['_raw_expiry'] == exp_date].drop(columns=['Symbol', '_raw_expiry'])
                    
                    if sub_df.empty:
                        continue

                    target_daily_date = sub_df['Daily'].iloc[0]

                    if not table_exists(cursor, table_name):
                        columns_def = [f'[{col}] TEXT' for col in sub_df.columns]
                        cursor.execute(f'CREATE TABLE [{table_name}] ({", ".join(columns_def)})')
                        conn.commit()

                    cursor.execute(f"SELECT COUNT(*) FROM [{table_name}] WHERE [Daily] = ?", (target_daily_date,))
                    if cursor.fetchone()[0] > 0:
                        continue

                    columns = ", ".join([f"[{col}]" for col in sub_df.columns])
                    placeholders = ", ".join(["?"] * len(sub_df.columns))
                    insert_sql = f"INSERT INTO [{table_name}] ({columns}) VALUES ({placeholders})"
                    
                    data_tuples = [tuple(x) for x in sub_df.to_numpy()]
                    cursor.executemany(insert_sql, data_tuples)
                    conn.commit()

            print(f"  -> Completed: {csv_file.name}")

        except Exception as exc:
            print(f"  [Error] Failed processing {csv_file.name}: {exc}")
            conn.rollback()

    cursor.close()
    conn.close()
    print('\nAll files processed successfully into MDB!')

if __name__ == "__main__":
    process_to_mdb()