"""Master Runner Script: Double-click to open terminal, choose options, and run the pipeline."""

from datetime import datetime, date, timedelta
from pathlib import Path

# Import functions from your script files
from Scraper import download_mcx_date_range
from Formater import clean_all_bhavcopies
from Converter import process_to_mdb

def main():
    print("=" * 60)
    print("      MCX DATA DOWNLOADER, FORMATTER & MDB UPDATER")
    print("=" * 60)
    print("\nSelect an option:")
    print("  1 - Download, format, and update with ONLY YESTERDAY'S data")
    print("  2 - Download, format, and update with a RANGE OF DAYS data")
    
    choice = input("\nEnter your choice (1 or 2): ").strip()
    
    start_date = None
    end_date = None

    if choice == "1":
        yesterday = date.today() - timedelta(days=1)
        start_date = yesterday
        end_date = yesterday
        print(f"\n[INFO] Selected Yesterday's Data: {yesterday.isoformat()}")

    elif choice == "2":
        print("\nEnter dates in DD/MM/YYYY format (e.g., 10/10/2026)")
        try:
            start_str = input("Enter Start Date: ").strip()
            end_str = input("Enter End Date: ").strip()
            
            start_date = datetime.strptime(start_str, "%d/%m/%Y").date()
            end_date = datetime.strptime(end_str, "%d/%m/%Y").date()
            
            if start_date > end_date:
                print("\n[ERROR] Start date cannot be after end date.")
                input("\nPress Enter to exit...")
                return
        except ValueError:
            print("\n[ERROR] Invalid date format! Please use DD/MM/YYYY.")
            input("\nPress Enter to exit...")
            return
    else:
        print("\n[ERROR] Invalid choice. Please restart the script and choose 1 or 2.")
        input("\nPress Enter to exit...")
        return

    print("\n" + "-" * 40)
    print("STEP 1/3: Running Scraper...")
    print("-" * 40)
    downloaded_raw_files = download_mcx_date_range(start_date, end_date, headful=True)

    if not downloaded_raw_files:
        print("\n[INFO] No new files were downloaded by the scraper. Pipeline finished early.")
        input("\nPress Enter to close this window...")
        return

    print("\n" + "-" * 40)
    print("STEP 2/3: Running Formatter...")
    print("-" * 40)
    formatted_csv_files = clean_all_bhavcopies(downloaded_raw_files)

    print("\n" + "-" * 40)
    print("STEP 3/3: Running MDB Converter...")
    print("-" * 40)
    process_to_mdb(formatted_csv_files)

    print("\n" + "=" * 60)
    print(" PIPELINE COMPLETED SUCCESSFULLY! ")
    print("=" * 60)
    input("\nPress Enter to close this window...")

if __name__ == "__main__":
    main()