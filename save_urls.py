
import os
import re
import time
import json
from io import StringIO
from urllib.parse import quote, urljoin
from difflib import SequenceMatcher
import pandas as pd
from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE = os.path.join(
    BASE_DIR,
    "NSE_BSE_Top_600s_Companies.xlsx"
)

INPUT_SHEET = "All Companies"

OUTPUT_DIR = os.path.join(BASE_DIR, "output")

OUTPUT_XLSX = os.path.join(
    OUTPUT_DIR,
    "Companies_urls.xlsx"
)

NSE_HOME = "https://www.nseindia.com"

NSE_ANNUAL_REPORT_URL = (
    "https://www.nseindia.com/"
    "companies-listing/corporate-filings-annual-reports"
)

NSE_EQUITY_CSV = (
    "https://nsearchives.nseindia.com/"
    "content/equities/EQUITY_L.csv"
)


BSE_HOME = "https://www.bseindia.com"

BSE_API_HOME = "https://api.bseindia.com"

BSE_SECURITY_MASTER_URL = (
    BSE_API_HOME +
    "/BseIndiaAPI/api/ListofScripData/w?segment=Equity"
)

# BSE corporate filing API.
BSE_ANNUAL_REPORT_API = (
    BSE_API_HOME +
    "/BseIndiaAPI/api/AnnSubCategoryGetData/w"
)

BSE_ATTACHMENT_BASE = (
    "https://www.bseindia.com/"
    "xml-data/corpfiling/AttachLive/"
)


# ============================================================
# TARGET FINANCIAL YEAR
# ============================================================

TARGET_FROM_YEAR = "2025"
TARGET_TO_YEAR = "2026"

HEADLESS = False

START_ROW = 1
END_ROW = None #None means process through the last company (all 600)

WAIT_AFTER_COMPANY = 2

TIMEOUT = 30000

MAX_RETRIES = 3

MATCH_THRESHOLD = 0.70

EXPECTED_COMPANIES = 600  


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/153.0.0.0 Safari/537.36"
)

NSE_HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": "text/html,application/json,*/*",
    "Referer": NSE_HOME
}

BSE_HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": "application/json,text/plain,*/*",
    "Referer": BSE_HOME,
    "Origin": BSE_HOME
}



def normalize_company_name(name):

    if pd.isna(name):
        return ""

    name = str(name).strip().lower()

    name = re.sub(
        r"[^a-z0-9]+",
        " ",
        name
    )

    suffixes = [
        "private limited",
        "private ltd",
        "pvt limited",
        "pvt ltd",
        "limited",
        "ltd",
        "incorporated",
        "inc",
        "corporation",
        "corp",
        "company",
        "co"
    ]

    for suffix in suffixes:

        name = re.sub(
            rf"\b{re.escape(suffix)}\b",
            " ",
            name
        )

    name = re.sub(
        r"\s+",
        " ",
        name
    ).strip()

    return name

def company_similarity(name1, name2):

    a = normalize_company_name(name1)
    b = normalize_company_name(name2)

    if not a or not b:
        return 0

    if a == b:
        return 1.0

    token_a = set(a.split())
    token_b = set(b.split())

    intersection = len(token_a & token_b)
    union = len(token_a | token_b)

    jaccard = intersection / union if union else 0

    sequence = SequenceMatcher(
        None,
        a,
        b
    ).ratio()

    return max(jaccard, sequence)


def load_input_file():

    print("\n" + "=" * 70)
    print("LOADING INPUT FILE")
    print("=" * 70)

    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(INPUT_FILE)

    df = pd.read_excel(
        INPUT_FILE,
        sheet_name=INPUT_SHEET,
        dtype=str,
        engine="openpyxl"
    ).fillna("")

    df.columns = df.columns.astype(str).str.strip()

    possible_columns = [
        "Company Name",
        "Company_Name",
        "Company",
        "company_name",
        "company"
    ]

    company_column = next(
        (
            c for c in possible_columns
            if c in df.columns
        ),
        None
    )

    if company_column is None:
        raise ValueError(
            "Company Name column not found. "
            f"Available columns: {list(df.columns)}"
        )

    result = pd.DataFrame()

    result["Company Name"] = (
        df[company_column]
        .astype(str)
        .str.strip()
    )

    result["Alternative pdf url"] = ""

    result = result[
        result["Company Name"] != ""
    ].reset_index(drop=True)

    result["_KEY"] = (
        result["Company Name"]
        .apply(normalize_company_name)
    )

    duplicates = result[
        result["_KEY"].duplicated(keep=False)
    ]

    if not duplicates.empty:

        raise ValueError(
            "Duplicate company names found:\n" +
            duplicates["Company Name"].to_string(
                index=False
            )
        )

    result = result.drop(columns=["_KEY"])

    print("Total companies:", len(result))

    if len(result) != EXPECTED_COMPANIES:

        raise ValueError(
            f"Expected {EXPECTED_COMPANIES} companies, "
            f"but found {len(result)}."
        )

    print("[PASSED] Exactly 600 companies loaded.")

    return result


def load_existing_output():

    if not os.path.exists(OUTPUT_XLSX):
        return None

    print("\nExisting output found.")
    print("Loading previous progress...")

    try:

        old_df = pd.read_excel(
            OUTPUT_XLSX,
            dtype=str,
            engine="openpyxl"
        ).fillna("")

        required = [
            "Company Name",
            "Alternative pdf url"
        ]

        if not all(c in old_df.columns for c in required):
            return None

        return old_df[required].copy()

    except Exception as e:

        print("Previous output error:", e)
        return None



def merge_previous_progress(current_df, old_df):

    if old_df is None:
        return current_df

    previous_results = {}

    for _, row in old_df.iterrows():

        company = str(
            row["Company Name"]
        ).strip()

        value = str(
            row["Alternative pdf url"]
        ).strip()

        if company:

            previous_results[
                normalize_company_name(company)
            ] = value

    for index in range(len(current_df)):

        company = current_df.at[
            index,
            "Company Name"
        ]

        key = normalize_company_name(company)

        if key in previous_results:

            current_df.at[
                index,
                "Alternative pdf url"
            ] = previous_results[key]

    return current_df


# ============================================================
# CHECK VALID RESULT
# ============================================================

def has_valid_existing_result(value):

    if pd.isna(value):
        return False

    value = str(value).strip()

    if not value:
        return False

    if value.lower() in [
        "nan",
        "none",
        "null"
    ]:
        return False

    if value.lower() == "not for this year":
        return True

    if value.startswith("http"):
        return True

    return False

def download_nse_equity_list(context):

    print("\nDownloading NSE equity master...")

    try:

        response = context.request.get(
            NSE_EQUITY_CSV,
            timeout=60000,
            headers=NSE_HEADERS
        )

        print("NSE status:", response.status)

        if not response.ok:
            return None

        text = response.body().decode(
            "utf-8-sig",
            errors="replace"
        )

        df = pd.read_csv(StringIO(text))

        print("NSE records:", len(df))

        return df

    except Exception as e:

        print("NSE master error:", e)
        return None

def find_nse_columns(df):

    symbol_column = None
    company_column = None

    for column in df.columns:

        normalized = str(column).strip().upper()

        if normalized == "SYMBOL":
            symbol_column = column

        if normalized in [
            "NAME OF COMPANY",
            "COMPANY NAME",
            "NAME_OF_COMPANY"
        ]:
            company_column = column

    return symbol_column, company_column


def build_nse_lookup(df):

    symbol_column, company_column = find_nse_columns(df)

    if symbol_column is None or company_column is None:
        raise ValueError("Required NSE columns not found.")

    lookup = []

    for _, row in df.iterrows():

        symbol = str(row[symbol_column]).strip().upper()
        company = str(row[company_column]).strip()

        if symbol and company:
            lookup.append({
                "company": company,
                "symbol": symbol
            })

    print("NSE company mappings:", len(lookup))

    return lookup


def find_nse_symbol(company_name, lookup):

    exact_key = normalize_company_name(company_name)

    # Exact match first.

    for item in lookup:

        if normalize_company_name(
            item["company"]
        ) == exact_key:

            return item["symbol"]

    # Similarity-based matching.

    best_symbol = None
    best_score = 0

    for item in lookup:

        score = company_similarity(
            company_name,
            item["company"]
        )

        if score > best_score:

            best_score = score
            best_symbol = item["symbol"]

    print(
        f"NSE best match score: {best_score:.3f}"
    )

    if best_score >= MATCH_THRESHOLD:
        return best_symbol

    return None


def open_nse_annual_report_page(page, symbol):

    url = (
        NSE_ANNUAL_REPORT_URL
        + "?symbol="
        + quote(symbol)
        + "&tabIndex=equity"
    )

    print("Opening NSE:", url)

    page.goto(
        url,
        wait_until="domcontentloaded",
        timeout=TIMEOUT
    )

    page.wait_for_timeout(4000)

    return url


def is_target_year_row(text):

    if not text:
        return False

    text = str(text)

    text = (
        text.replace("–", "-")
        .replace("—", "-")
        .replace("/", "-")
    )

    has_from = bool(
        re.search(
            rf"\b{TARGET_FROM_YEAR}\b",
            text
        )
    )

    has_to = bool(
        re.search(
            rf"\b{TARGET_TO_YEAR}\b",
            text
        )
    )

    return has_from and has_to

def clean_url(url):

    if not url:
        return ""

    url = str(url).strip()

    url = url.replace("&amp;", "&")

    url = url.strip("\"'()[]{}")

    return url


def extract_pdf_url_from_element(element):

    attributes = [
        "href",
        "data-url",
        "data-href",
        "data-link",
        "data-pdf",
        "data-pdf-url",
        "data-download"
    ]

    for attribute in attributes:

        try:

            value = element.get_attribute(attribute)

            if not value:
                continue

            value = clean_url(value)

            if (
                ".pdf" in value.lower()
                or "nsearchives.nseindia.com" in value.lower()
            ):

                return urljoin(
                    NSE_HOME,
                    value
                )

        except Exception:
            continue

    try:

        onclick = element.get_attribute("onclick") or ""

        matches = re.findall(
            r"""https?://[^"'()\s]+""",
            onclick,
            re.IGNORECASE
        )

        for value in matches:

            value = clean_url(value)

            if ".pdf" in value.lower():
                return value

    except Exception:
        pass

    return ""

def extract_nse_url_from_row(row):

    try:

        links = row.locator("a")

        for i in range(links.count()):

            url = extract_pdf_url_from_element(
                links.nth(i)
            )

            if url:
                return url

        buttons = row.locator("button")

        for i in range(buttons.count()):

            url = extract_pdf_url_from_element(
                buttons.nth(i)
            )

            if url:
                return url

        html = row.evaluate(
            "(el) => el.outerHTML"
        )

        matches = re.findall(
            r"""https?://[^"'<> ]+?\.pdf(?:\?[^"'<> ]*)?""",
            html,
            re.IGNORECASE
        )

        if matches:
            return clean_url(matches[0])

    except Exception as e:

        print("NSE row extraction error:", e)

    return ""

def find_nse_report(page):

    print(f"\nSearching NSE for FY {TARGET_FROM_YEAR}-{TARGET_TO_YEAR} report...")

    page.wait_for_timeout(3000)

    rows = page.locator("table tbody tr")

    try:
        count = rows.count()
    except Exception:
        count = 0

    print("NSE table rows:", count)

    target_found = False

    for i in range(count):

        row = rows.nth(i)

        try:

            if not row.is_visible():
                continue

            text = " ".join(
                row.inner_text().split()
            )

            if not is_target_year_row(text):
                continue

            target_found = True

            print("NSE target row:", text[:300])

            url = extract_nse_url_from_row(row)

            if url:
                return url

        except Exception:
            continue

    if target_found:

        print("NSE report row found, but PDF URL unavailable.")

        return ""

    print(f"NSE FY {TARGET_FROM_YEAR}-{TARGET_TO_YEAR} report not found.")

    return None

def download_bse_security_master(context):

    print("\nDownloading BSE security master...")

    try:

        response = context.request.get(
            BSE_SECURITY_MASTER_URL,
            timeout=60000,
            headers=BSE_HEADERS
        )

        print("BSE master status:", response.status)

        if not response.ok:
            return None

        body = response.body()

        content_type = response.headers.get(
            "content-type",
            ""
        ).lower()

        if "json" in content_type or body.lstrip().startswith(
            (b"{", b"[")
        ):

            data = json.loads(
                body.decode("utf-8-sig")
            )

            if isinstance(data, dict):

                for key in [
                    "Table",
                    "table",
                    "data",
                    "Data",
                    "result"
                ]:

                    if isinstance(data.get(key), list):
                        data = data[key]
                        break

            if isinstance(data, list):
                return pd.DataFrame(data)

        # Some BSE responses may be CSV.

        text = body.decode(
            "utf-8-sig",
            errors="replace"
        )

        return pd.read_csv(StringIO(text))

    except Exception as e:

        print("BSE master error:", e)
        return None


def find_bse_columns(df):

    code_column = None
    company_column = None

    possible_codes = [
        "SCRIP_CD",
        "SCRIPCODE",
        "SCRIP_CODE",
        "SECURITY_CODE",
        "SC_CODE",
        "SCRIP CD",
        "CODE"
    ]

    possible_names = [
        "SCRIP_NAME",
        "SCRIPNAME",
        "SECURITY_NAME",
        "COMPANY_NAME",
        "ISSUER_NAME",
        "NAME",
        "SCRIP NAME"
    ]

    for column in df.columns:

        normalized = (
            str(column)
            .strip()
            .upper()
            .replace("-", "_")
        )

        if normalized in possible_codes:
            code_column = column

        if normalized in possible_names:
            company_column = column

    return code_column, company_column

def build_bse_lookup(df):

    code_column, company_column = find_bse_columns(df)

    if code_column is None or company_column is None:

        print("BSE master columns:")
        print(list(df.columns))

        raise ValueError(
            "Could not identify BSE security-code "
            "and company-name columns. Check the "
            "current BSE master response."
        )

    lookup = []

    for _, row in df.iterrows():

        company = str(
            row[company_column]
        ).strip()

        code = str(
            row[code_column]
        ).strip()

        if not company or not code:
            continue

        if code.endswith(".0"):
            code = code[:-2]

        lookup.append({
            "company": company,
            "code": code
        })

    print("BSE company mappings:", len(lookup))

    return lookup

def find_bse_code(company_name, lookup):

    target = normalize_company_name(company_name)

    # Exact match first.

    for item in lookup:

        if normalize_company_name(
            item["company"]
        ) == target:

            print("Exact BSE match:", item)

            return item["code"]

    # Fuzzy match.

    best_code = None
    best_company = None
    best_score = 0

    for item in lookup:

        score = company_similarity(
            company_name,
            item["company"]
        )

        if score > best_score:

            best_score = score
            best_code = item["code"]
            best_company = item["company"]

    print(
        f"BSE best match: {best_company} "
        f"(score={best_score:.3f})"
    )

    if best_score >= MATCH_THRESHOLD:
        return best_code

    return None


def find_attachment_values(data):

    results = []

    def walk(obj):

        if isinstance(obj, dict):

            for key, value in obj.items():

                key_normalized = str(key).lower()

                if (
                    "attachment" in key_normalized
                    or "pdf" in key_normalized
                    or "file" in key_normalized
                ):

                    if isinstance(value, str) and value.strip():
                        results.append(value.strip())

                if isinstance(value, (dict, list)):
                    walk(value)

        elif isinstance(obj, list):

            for item in obj:
                walk(item)

    walk(data)

    return list(dict.fromkeys(results))


def search_bse_annual_report(
    context,
    company_name,
    security_code
):

    print("\nSearching BSE annual reports...")

    print("Company:", company_name)
    print("BSE code:", security_code)

    today = time.strftime("%Y-%m-%d")

    params = {
        "strCat": "Company Update",
        "strScrip": str(security_code),
        "strSearch": "P",
        "strToDate": today,
        "strType": "C",
        "subcategory": "Annual Report"
    }

    try:

        response = context.request.get(
            BSE_ANNUAL_REPORT_API,
            params=params,
            timeout=60000,
            headers=BSE_HEADERS
        )

        print("BSE filing API status:", response.status)

        if not response.ok:
            return None

        data = response.json()

        if isinstance(data, dict):

            records = None

            for key in [
                "Table",
                "table",
                "data",
                "Data",
                "result"
            ]:

                if isinstance(data.get(key), list):
                    records = data[key]
                    break

            if records is None:
                records = [data]

        elif isinstance(data, list):

            records = data

        else:
            return None

        for record in records:

            if not isinstance(record, dict):
                continue

            record_text = json.dumps(
                record,
                ensure_ascii=False
            )

            # Require both financial years in the filing
            # metadata or description.

            if not is_target_year_row(record_text):
                continue

            print(f"BSE FY {TARGET_FROM_YEAR}-{TARGET_TO_YEAR} filing found.")

            attachments = find_attachment_values(record)

            for attachment in attachments:

                attachment = clean_url(attachment)

                if attachment.lower().endswith(".pdf"):

                    if attachment.startswith("http"):
                        return attachment

                    return urljoin(
                        BSE_ATTACHMENT_BASE,
                        attachment.lstrip("/")
                    )

                # Some responses provide an attachment
                # identifier without a .pdf extension.

                if (
                    "/" not in attachment
                    and "\\" not in attachment
                    and len(attachment) > 3
                ):

                    return urljoin(
                        BSE_ATTACHMENT_BASE,
                        attachment
                    )

        print(f"No matching BSE FY {TARGET_FROM_YEAR}-{TARGET_TO_YEAR} filing.")

        return None

    except Exception as e:

        print("BSE annual report API error:", e)

        return None


def verify_pdf_url(context, url):

    if not url:
        return False

    try:

        response = context.request.get(
            url,
            timeout=60000,
            headers={
                **BSE_HEADERS,
                "Range": "bytes=0-1023"
            }
        )

        if not response.ok:
            return False

        content_type = response.headers.get(
            "content-type",
            ""
        ).lower()

        body = response.body()

        is_pdf = (
            "application/pdf" in content_type
            or body.startswith(b"%PDF")
        )

        return is_pdf

    except Exception as e:

        print("PDF verification error:", e)

        return False


# ============================================================
# PROCESS NSE FIRST, THEN BSE FALLBACK
# ============================================================

def process_company(
    page,
    context,
    company_name,
    nse_lookup,
    bse_lookup
):

    print("\n" + "=" * 75)
    print("COMPANY:", company_name)
    print("=" * 75)

    # --------------------------------------------------------
    # STEP 1: NSE SEARCH
    # --------------------------------------------------------

    print("\nSTEP 1: NSE SEARCH")

    nse_symbol = find_nse_symbol(
        company_name,
        nse_lookup
    )

    if nse_symbol:

        print("NSE symbol:", nse_symbol)

        for attempt in range(1, MAX_RETRIES + 1):

            print(
                f"NSE attempt {attempt}/{MAX_RETRIES}"
            )

            try:

                open_nse_annual_report_page(
                    page,
                    nse_symbol
                )

                result = find_nse_report(page)

                if result:

                    if verify_pdf_url(context, result):

                        print("[SUCCESS] Verified NSE PDF.")

                        return result

                    print("NSE PDF URL failed verification.")

                if result == "":
                    print("NSE row found, URL extraction failed.")

            except Exception as e:

                print("NSE processing error:", e)

            if attempt < MAX_RETRIES:
                time.sleep(3)

    else:

        print("Company not matched in NSE.")

    # --------------------------------------------------------
    # STEP 2: BSE FALLBACK
    # --------------------------------------------------------

    print("\nSTEP 2: BSE FALLBACK")

    bse_code = find_bse_code(
        company_name,
        bse_lookup
    )

    if not bse_code:

        print("Company not matched in BSE.")

        return ""

    print("BSE security code:", bse_code)

    for attempt in range(1, MAX_RETRIES + 1):

        print(
            f"BSE attempt {attempt}/{MAX_RETRIES}"
        )

        result = search_bse_annual_report(
            context,
            company_name,
            bse_code
        )

        if result:

            if verify_pdf_url(context, result):

                print("[SUCCESS] Verified BSE PDF.")

                return result

            print("BSE PDF failed verification.")

        if attempt < MAX_RETRIES:
            time.sleep(3)

    print("No verified report obtained from NSE or BSE.")

    return ""


def save_files(df):

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )

    final_df = df[
        [
            "Company Name",
            "Alternative pdf url"
        ]
    ].copy()

    temp_file = OUTPUT_XLSX + ".tmp.xlsx"

    with pd.ExcelWriter(
        temp_file,
        engine="openpyxl"
    ) as writer:

        final_df.to_excel(
            writer,
            sheet_name="Company URLs",
            index=False
        )

        worksheet = writer.sheets["Company URLs"]

        worksheet.freeze_panes = "A2"

        worksheet.column_dimensions["A"].width = 45

        worksheet.column_dimensions["B"].width = 100

        worksheet.auto_filter.ref = worksheet.dimensions

    os.replace(
        temp_file,
        OUTPUT_XLSX
    )

    print("\nProgress saved:", OUTPUT_XLSX)


def show_progress(df):

    total = len(df)

    completed = sum(
        has_valid_existing_result(value)
        for value in df["Alternative pdf url"]
    )

    remaining = total - completed

    print("\n" + "-" * 70)

    print("TOTAL     :", total)
    print("COMPLETED :", completed)
    print("REMAINING :", remaining)

    print("-" * 70)



def main():

    print("\n" + "=" * 75)

    print(f"NSE + BSE FY {TARGET_FROM_YEAR}-{TARGET_TO_YEAR} ANNUAL REPORT COLLECTOR")

    print("=" * 75)
    df = load_input_file()
    old_df = load_existing_output()

    if old_df is not None:

        df = merge_previous_progress(
            df,
            old_df
        )

        print("\nPrevious progress restored.")

    show_progress(df)

    start_index = max(
        0,
        START_ROW - 1
    )

    if END_ROW is None:

        end_index = len(df)

    else:

        end_index = min(
            END_ROW,
            len(df)
        )

    print(
        f"\nProcessing rows {start_index + 1} "
        f"to {end_index}"
    )


    with sync_playwright() as p:

        print("\nStarting Chromium...")

        browser = p.chromium.launch(
            headless=HEADLESS
        )

        try:

            context = browser.new_context(

                accept_downloads=False,

                viewport={
                    "width": 1440,
                    "height": 1000
                },

                user_agent=USER_AGENT,

                locale="en-IN"
            )

            page = context.new_page()

            page.set_default_timeout(TIMEOUT)

            print("\nOpening NSE...")

            try:

                page.goto(
                    NSE_HOME,
                    wait_until="domcontentloaded",
                    timeout=TIMEOUT
                )

                page.wait_for_timeout(3000)

            except Exception as e:

                print("NSE opening error:", e)
            nse_df = download_nse_equity_list(context)

            if nse_df is None:

                raise RuntimeError(
                    "NSE master unavailable. "
                    "Cannot safely start the processing."
                )

            nse_lookup = build_nse_lookup(nse_df)

            bse_df = download_bse_security_master(context)

            if bse_df is None:

                print(
                    "\nWARNING: BSE master unavailable. "
                    "BSE fallback will not work until "
                    "the BSE master endpoint is corrected."
                )

                bse_lookup = []

            else:

                try:

                    bse_lookup = build_bse_lookup(bse_df)

                except Exception as e:

                    print("BSE lookup error:", e)

                    bse_lookup = []

            for index in range(
                start_index,
                end_index
            ):

                company_name = str(
                    df.at[index, "Company Name"]
                ).strip()

                if not company_name:
                    continue

                existing_value = str(
                    df.at[
                        index,
                        "Alternative pdf url"
                    ]
                ).strip()

           

                if has_valid_existing_result(existing_value):

                    print("\n" + "-" * 75)

                    print(
                        f"[{index + 1}/{len(df)}] "
                        f"{company_name}"
                    )

                    print("Already completed. Skipping.")

                    continue

                print("\n" + "#" * 75)

                print(
                    f"[{index + 1}/{len(df)}] "
                    f"{company_name}"
                )

                print("#" * 75)
                result = process_company(
                    page,
                    context,
                    company_name,
                    nse_lookup,
                    bse_lookup
                )

                df.at[
                    index,
                    "Alternative pdf url"
                ] = result


                save_files(df)

                if result:

                    print("\nFINAL RESULT:")
                    print(result)

                else:

                    print(
                        "\nNo verified URL. "
                        "Will retry on next execution."
                    )

                time.sleep(WAIT_AFTER_COMPANY)

        finally:

            browser.close()

    save_files(df)

    print("\n" + "=" * 75)

    print("PROCESS COMPLETED")

    print("=" * 75)

    show_progress(df)

    print("\nFINAL OUTPUT:")
    print(OUTPUT_XLSX)

    print("\nOUTPUT COLUMNS:")
    print("1. Company Name")
    print("2. Alternative pdf url")

if __name__ == "__main__":

    main()
