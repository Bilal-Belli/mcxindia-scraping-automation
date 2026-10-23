
"""Master runner: scrape, format, merge, and synchronize the MDB."""

from datetime import datetime, date, timedelta

from Scraper import download_mcx_date_range
from Formater import clean_all_bhavcopies
from Converter import process_to_mdb


def main():
    print("=" * 60)
    print("      MCX DATA DOWNLOADER, FORMATTER & MDB UPDATER")
    print("=" * 60)

    print("\nSelect an option:")
    print("  1 - Download and update yesterday's data")
    print("  2 - Download and update a date range")

    choice = input("\nEnter your choice (1 or 2): ").strip()

    if choice == "1":
        start_date = end_date = date.today() - timedelta(days=1)

    elif choice == "2":
        print("\nEnter dates in DD/MM/YYYY format.")

        try:
            start_date = datetime.strptime(
                input("Start date: ").strip(), "%d/%m/%Y"
            ).date()

            end_date = datetime.strptime(
                input("End date: ").strip(), "%d/%m/%Y"
            ).date()

            if start_date > end_date:
                raise ValueError("Start date must not exceed end date.")

        except ValueError as exc:
            print(f"\n[ERROR] {exc}")
            input("\nPress Enter to exit...")
            return

    else:
        print("\n[ERROR] Invalid choice.")
        input("\nPress Enter to exit...")
        return

    print(
        f"\nSelected period: {start_date:%d/%m/%Y} "
        f"to {end_date:%d/%m/%Y}"
    )

    try:
        # STEP 1: Download missing raw files.
        print("\n" + "-" * 50)
        print("STEP 1/3: SCRAPER")
        print("-" * 50)

        downloaded = download_mcx_date_range(
            start_date,
            end_date,
            headful=True,
        )

        print(f"New raw downloads: {len(downloaded)}")

        # STEP 2: Format new raw files.
        print("\n" + "-" * 50)
        print("STEP 2/3: FORMATTER")
        print("-" * 50)

        if downloaded:
            formatted = clean_all_bhavcopies(downloaded)
        else:
            formatted = clean_all_bhavcopies()

        print(f"New formatted files: {len(formatted)}")

        # STEP 3: Merge into master CSV and synchronize Access.
        print("\n" + "-" * 50)
        print("STEP 3/3: MASTER CSV + MDB")
        print("-" * 50)

        if formatted:
            # Merge newly formatted files into the master CSV.
            process_to_mdb(formatted)
        else:
            # Recover from existing filtered files or synchronize
            # the existing master CSV.
            process_to_mdb()

        print("\n" + "=" * 60)
        print("PIPELINE COMPLETED")
        print("=" * 60)

    except Exception as exc:
        print(f"\n[ERROR] Pipeline failed: {type(exc).__name__}: {exc}")
        print("Correct the error and rerun the pipeline.")


if __name__ == "__main__":
    main()