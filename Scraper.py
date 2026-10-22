"""Standalone script to download MCX Bhavcopy CSVs for all days within a date range into a local 'bhavcopies' folder."""

from datetime import date, timedelta
from pathlib import Path
from playwright.sync_api import sync_playwright

PAGE_URL = "https://www.mcxindia.com/market-data/bhavcopy"
OUTPUT_DIR = Path(__file__).resolve().parent / "bhavcopies"

def download_mcx_date_range(start_date: date, end_date: date, headful: bool = True) -> list[Path]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    current_date = start_date
    downloaded_files = []
    print(f"Starting fast batch download for all days from {start_date.isoformat()} to {end_date.isoformat()}...\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=not headful,
            args=["--disable-blink-features=AutomationControlled"]
        )
        context = browser.new_context(
            accept_downloads=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        try:
            print("Navigating to MCX Bhavcopy page...")
            page.goto(PAGE_URL, timeout=30_000)
        except Exception as exc:
            print(f"Failed to load page: {exc}")
            browser.close()
            return downloaded_files

        while current_date <= end_date:
            date_str_ddmmyyyy = current_date.strftime("%d%m%Y")
            file_name = f"BhavCopyDateWise_{date_str_ddmmyyyy}.csv"
            file_path = OUTPUT_DIR / file_name

            if file_path.exists():
                print(f"[{current_date.isoformat()}] File exists. Skipping.")
                current_date += timedelta(days=1)
                continue

            date_formatted = current_date.strftime("%d/%m/%Y")

            try:
                page.evaluate(f"""() => {{
                    let input = document.getElementById('bhavCopyDateWise');
                    if (input) {{
                        input.value = '{date_formatted}';
                        input.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    }}
                }}""")

                page.locator("#btnDateWise").click()
                
                try:
                    page.wait_for_selector("#btnDateWiseCSV", state="visible", timeout=2_500)
                    csv_button = page.locator("#btnDateWiseCSV")

                    with page.expect_download(timeout=5_000) as download_info:
                        csv_button.click()
                    
                    download_info.value.save_as(file_path)
                    print(f"[{date_formatted}] Downloaded: {file_name}")
                    downloaded_files.append(file_path)
                except Exception:
                    print(f"[{date_formatted}] No data / Holiday. Skipping quickly.")

            except Exception as exc:
                print(f"[{date_formatted}] Error: {exc}")

            current_date += timedelta(days=1)

        browser.close()
        print(f"\nBatch download finished! {len(downloaded_files)} new file(s) downloaded.")
        return downloaded_files

if __name__ == "__main__":
    start_d = date.today() - timedelta(days=14)
    download_mcx_date_range(start_d, date.today(), headful=False)