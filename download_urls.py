import os
import re
import time
import requests
import pandas as pd

from pathlib import Path
from urllib.parse import urlparse

# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = BASE_DIR / "output" / "Companies_urls.xlsx"

URL_COLUMN = "Alternative pdf url"
COMPANY_COLUMN = "Company Name"
# ============================================================
# COMPANY PROCESSING RANGE
# ============================================================

START_ROW = 1
END_ROW = 600

# Set END_ROW = None to process through the last company.
PDF_FOLDER = BASE_DIR / "annual_reports"

# Request settings
REQUEST_TIMEOUT = 60
MAX_RETRIES = 3
RETRY_DELAY = 3
MIN_PDF_SIZE = 10_000

# ============================================================
# CREATE PDF FOLDER
# ============================================================

PDF_FOLDER.mkdir(parents=True, exist_ok=True)


# ============================================================
# SAFE COMPANY NAME
# ============================================================

def safe_filename(company_name):

    name = str(company_name).strip()

    # Remove characters not allowed in Windows filenames
    name = re.sub(r'[<>:"/\\|?*]', '', name)

    # Remove extra spaces
    name = re.sub(r'\s+', ' ', name)

    # Remove trailing dots and spaces
    name = name.rstrip(". ")

    return name.strip()


# ============================================================
# CHECK IF ALREADY DOWNLOADED
# ============================================================

def already_downloaded(company_name):

    filename = safe_filename(company_name) + ".pdf"

    path = PDF_FOLDER / filename

    return (
        path.exists()
        and path.stat().st_size > MIN_PDF_SIZE
    )


# ============================================================
# LOAD INPUT FILE
# ============================================================

def load_input_file():

    if not os.path.exists(INPUT_FILE):

        raise FileNotFoundError(
            f"Input file not found: {INPUT_FILE}"
        )

    file_extension = Path(INPUT_FILE).suffix.lower()

    print(f"\nReading input file: {INPUT_FILE}")

    if file_extension in [".xlsx", ".xls"]:

        df = pd.read_excel(
            INPUT_FILE,
            engine="openpyxl" if file_extension == ".xlsx" else None
        )

    elif file_extension == ".csv":

        try:
            df = pd.read_csv(
                INPUT_FILE,
                encoding="utf-8-sig"
            )

        except UnicodeDecodeError:

            df = pd.read_csv(
                INPUT_FILE,
                encoding="cp1252"
            )

    else:

        raise ValueError(
            f"Unsupported file format: {file_extension}"
        )

    # Clean column names
    df.columns = (
        df.columns.astype(str).str.strip()
    )

    print("\nInput file loaded successfully.")

    print("Columns:", df.columns.tolist())

    # Validate required columns
    if COMPANY_COLUMN not in df.columns:

        raise ValueError(
            f"Missing company column: {COMPANY_COLUMN}"
        )

    if URL_COLUMN not in df.columns:

        raise ValueError(
            f"Missing URL column: {URL_COLUMN}"
        )

    # Remove rows without company names
    df = df.dropna(
        subset=[COMPANY_COLUMN]
    ).copy()

    return df


# ============================================================
# DOWNLOAD ONE PDF
# ============================================================

def download_pdf(company_name, pdf_url):

    filename = safe_filename(company_name) + ".pdf"

    output_path = PDF_FOLDER / filename

    # --------------------------------------------------------
    # Skip if already downloaded
    # --------------------------------------------------------

    if already_downloaded(company_name):

        print("   SKIP - Already downloaded")
        print(f"   {output_path}")

        return "skipped"

    # --------------------------------------------------------
    # Validate URL
    # --------------------------------------------------------

    if pd.isna(pdf_url):

        print("   No PDF URL")

        return "no_url"

    pdf_url = str(pdf_url).strip()

    if pdf_url.lower() in [
        "",
        "nan",
        "none",
        "null",
        "not for this year",
        "not found"
    ]:

        print("   No valid PDF URL")

        return "no_url"

    # Validate URL format
    parsed_url = urlparse(pdf_url)

    if parsed_url.scheme not in ["http", "https"] or not parsed_url.netloc:

        print("   Invalid URL format")

        return "no_url"

    print(f"   URL: {pdf_url}")

    # --------------------------------------------------------
    # Headers
    # --------------------------------------------------------

    headers = {

        "User-Agent":
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/131.0.0.0 Safari/537.36",

        "Accept":
            "application/pdf,*/*;q=0.8",

        "Referer":
            "https://www.nseindia.com/"
    }

    # --------------------------------------------------------
    # Retry
    # --------------------------------------------------------

    for attempt in range(1, MAX_RETRIES + 1):

        try:

            print(
                f"   Download attempt {attempt}/{MAX_RETRIES}"
            )

            response = requests.get(
                pdf_url,
                headers=headers,
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True
            )

            print(
                f"   HTTP status: {response.status_code}"
            )

            response.raise_for_status()

            content = response.content

            # ------------------------------------------------
            # Validate PDF size
            # ------------------------------------------------

            if len(content) < MIN_PDF_SIZE:

                raise ValueError(
                    f"Downloaded content too small: {len(content)} bytes"
                )

            # ------------------------------------------------
            # Validate PDF signature
            # ------------------------------------------------

            if not content.lstrip().startswith(b"%PDF"):

                content_type = response.headers.get(
                    "Content-Type", ""
                ).lower()

                raise ValueError(
                    f"Response is not a valid PDF. "
                    f"Content-Type: {content_type}"
                )

            # ------------------------------------------------
            # Save PDF safely
            # ------------------------------------------------

            temp_path = output_path.with_suffix(".tmp")

            temp_path.write_bytes(content)

            # Rename only after successful writing
            temp_path.replace(output_path)

            print("   SUCCESS")

            print(f"   Saved: {output_path}")

            print(
                f"   Size: {len(content) / (1024 * 1024):.2f} MB"
            )

            return "downloaded"

        except Exception as e:

            print(f"   ERROR: {e}")

            if attempt < MAX_RETRIES:

                print(
                    f"   Retrying in {RETRY_DELAY} seconds..."
                )

                time.sleep(RETRY_DELAY)

    print("   FAILED after all retries.")

    return "failed"


# ============================================================
# MAIN
# ============================================================


def main():

    print("\n")
    print("=" * 70)
    print("NSE/BSE ANNUAL REPORT PDF DOWNLOADER")
    print("=" * 70)

    # --------------------------------------------------------
    # LOAD INPUT FILE
    # --------------------------------------------------------

    df = load_input_file()

    total_companies = len(df)

    print(f"\nTotal companies in input file: {total_companies}")

    print(f"PDF folder: {PDF_FOLDER.absolute()}")

    # --------------------------------------------------------
    # VALIDATE START AND END ROW
    # --------------------------------------------------------

    if START_ROW < 1:
        raise ValueError("START_ROW must be at least 1.")

    if END_ROW is not None and END_ROW < START_ROW:
        raise ValueError("END_ROW cannot be smaller than START_ROW.")

    if START_ROW > total_companies:
        raise ValueError(
            f"START_ROW ({START_ROW}) exceeds the total "
            f"number of companies ({total_companies})."
        )

    start_index = START_ROW - 1

    if END_ROW is None:
        end_index = total_companies
    else:
        end_index = min(END_ROW, total_companies)

    selected_df = df.iloc[start_index:end_index].copy()

    print("\n" + "=" * 70)
    print("COMPANY PROCESSING RANGE")
    print("=" * 70)

    print(f"Starting company row : {START_ROW}")
    print(f"Ending company row   : {end_index}")
    print(f"Companies selected   : {len(selected_df)}")

    print("=" * 70)

    if selected_df.empty:
        raise RuntimeError("No companies selected for downloading.")

    # --------------------------------------------------------
    # COUNTERS
    # --------------------------------------------------------

    downloaded = 0
    skipped = 0
    failed = 0
    no_url = 0

    # --------------------------------------------------------
    # PROCESS SELECTED COMPANIES
    # --------------------------------------------------------

    for position, (index, row) in enumerate(
        selected_df.iterrows(),
        start=1
    ):

        company_name = str(
            row[COMPANY_COLUMN]
        ).strip()

        pdf_url = row[URL_COLUMN]

        original_row_number = index + 1

        print("\n")
        print("=" * 70)

        print(
            f"[{original_row_number}/{total_companies}] "
            f"{company_name}"
        )

        print(
            f"Selected batch progress: "
            f"{position}/{len(selected_df)}"
        )

        print("=" * 70)

        # ----------------------------------------------------
        # EMPTY COMPANY
        # ----------------------------------------------------

        if not company_name or company_name.lower() == "nan":

            print("   Invalid company name")

            failed += 1

            continue

        # ----------------------------------------------------
        # DOWNLOAD PDF
        # ----------------------------------------------------

        result = download_pdf(
            company_name,
            pdf_url
        )

        if result == "downloaded":
            downloaded += 1

        elif result == "skipped":
            skipped += 1

        elif result == "no_url":
            no_url += 1

        else:
            failed += 1

        # Small delay between companies
        time.sleep(1)

    # ========================================================
    # FINAL SUMMARY
    # ========================================================

    processed = len(selected_df)

    print("\n")
    print("=" * 70)
    print("DOWNLOAD COMPLETE")
    print("=" * 70)

    print(f"Total companies in Excel : {total_companies}")
    print(f"Selected start row       : {START_ROW}")
    print(f"Selected end row         : {end_index}")
    print(f"Companies processed      : {processed}")

    print("-" * 70)

    print(f"Downloaded               : {downloaded}")
    print(f"Already existed           : {skipped}")
    print(f"No URL                    : {no_url}")
    print(f"Failed                    : {failed}")

    print("=" * 70)

    print("\nPDF folder:")
    print(PDF_FOLDER.absolute())

    print("=" * 70)



# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    main()
