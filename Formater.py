"""Standalone script to clean all MCX Bhavcopy CSV files, filter them, and save into a 'filtredData' folder."""

from datetime import datetime
from pathlib import Path
import pandas as pd

INPUT_DIR = Path(__file__).resolve().parent / "bhavcopies"
OUTPUT_DIR = Path(__file__).resolve().parent / "filtredData"

def clean_all_bhavcopies(input_files: list[Path] = None) -> list[Path]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    if input_files is None:
        files = sorted(INPUT_DIR.glob("BhavCopyDateWise_*.csv"))
    else:
        files = sorted([f for f in input_files if f.exists()])

    if not files:
        print(f"No BhavCopyDateWise CSV files to process.")
        return []
    
    print(f"Found {len(files)} file(s) total.\n")
    formatted_files = []

    for input_file in files:
        date_str = None
        
        # Try to extract the date from the filename first (avoids loading CSV if already exists)
        try:
            filename_part = input_file.stem.split('_')[1] 
            if len(filename_part) == 8 and filename_part.isdigit():
                date_str = filename_part
        except (IndexError, ValueError):
            pass

        # If date found from filename, check if output already exists
        if date_str:
            output_filename = f"MCX{date_str}.csv"
            output_path = OUTPUT_DIR / output_filename
            if output_path.exists():
                print(f"Skipping {input_file.name} (already formatted: {output_path.name})")
                continue

        print(f"Processing raw file: {input_file.name}")

        try:
            df = pd.read_csv(input_file, dtype=str, keep_default_na=False, encoding="utf-8-sig")
            df.columns = df.columns.str.strip()

            if 'Option Type' in df.columns:
                df = df[df['Option Type'].str.strip() == '-']

            columns_to_drop = [
                'Instrument Name', 
                'Option Type', 
                'Strike Price', 
                "Volume(In 000's)"
            ]
            existing_cols_to_drop = [col for col in columns_to_drop if col in df.columns]
            df = df.drop(columns=existing_cols_to_drop)

            # Fallback to dataframe date if filename date wasn't parsed correctly
            if not date_str:
                try:
                    if 'Date' in df.columns and not df.empty:
                        sample_date_text = df['Date'].iloc[0].strip()
                        parsed_date = datetime.strptime(sample_date_text, "%d %b %Y")
                        date_str = parsed_date.strftime("%d%m%Y")
                    else:
                        raise ValueError()
                except Exception:
                    print(f"  [Warning] Could not determine date for {input_file.name}. Skipping.")
                    continue

                output_filename = f"MCX{date_str}.csv"
                output_path = OUTPUT_DIR / output_filename
                if output_path.exists():
                    print(f"  -> Skipping (already formatted: {output_path.name})\n")
                    continue

            df.to_csv(output_path, index=False)
            print(f"  -> Successfully saved to: {output_path.name}\n")
            formatted_files.append(output_path)

        except Exception as exc:
            print(f"  [Error] Failed to process {input_file.name}: {exc}\n")

    print("Batch cleaning process finished!")
    return formatted_files

if __name__ == "__main__":
    clean_all_bhavcopies()