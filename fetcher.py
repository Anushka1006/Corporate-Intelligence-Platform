import re
import sys
import time
from io import BytesIO, StringIO
from pathlib import Path
import pandas as pd
import requests
BASE_DIR = Path(__file__).resolve().parent

OUTPUT_FILE = BASE_DIR / "NSE_BSE_Top_600s_Companies.xlsx"
DEBUG_DIR = BASE_DIR / "bse_debug"

TARGET_TOTAL = 600
EXPECTED_NSE_COUNT = 100

NSE_TOP100_URL = (
    "https://www.niftyindices.com/"
    "IndexConstituent/ind_nifty100list.csv"
)

NSE_EQUITY_MASTER_URL = (
    "https://nsearchives.nseindia.com/"
    "content/equities/EQUITY_L.csv"
)

NSE_HOME_URL = "https://www.nseindia.com"

BSE_TOP500_API_URL = (
    "https://www.bseindices.com/AsiaIndexAPI/"
    "api/Codewise_IndicesDownload/w?code=17"
)

REQUEST_TIMEOUT = 60
MAX_RETRIES = 3

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/130.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,"
        "application/xml;q=0.9,*/*;q=0.8"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

def normalize_company_name(name):
    """
    Normalize company names for matching NSE and BSE records.
    """

    if pd.isna(name):
        return ""

    name = str(name).strip().upper()

    if not name or name in {"NAN", "NONE", "NULL"}:
        return ""

    name = name.replace("&", " AND ")
    name = re.sub(r"[^A-Z0-9 ]", " ", name)
    name = re.sub(r"\s+", " ", name).strip()

    suffixes = [
        " LIMITED",
        " LTD",
        " LTD.",
        " PRIVATE LIMITED",
        " PVT LIMITED",
        " PVT LTD",
        " PVT. LTD.",
        " PRIVATE LTD",
        " PRIVATE",
        " INCORPORATED",
        " INC",
        " CORPORATION",
        " CORP",
        " COMPANY",
        " CO",
    ]

    changed = True

    while changed:
        changed = False

        for suffix in suffixes:
            if name.endswith(suffix):
                name = name[:-len(suffix)].strip()
                changed = True

    return name


def clean_isin(value):
    """
    Clean ISIN values.
    """

    if pd.isna(value):
        return ""

    value = str(value).strip().upper()

    if value in {"", "NAN", "NONE", "NULL", "-"}:
        return ""

    return value


def clean_text(value):
    """
    Convert a value into a clean string.
    """

    if pd.isna(value):
        return ""

    value = str(value).strip()

    if value.lower() in {"nan", "none", "null"}:
        return ""

    return value


def find_column(df, possible_names):
    """
    Find a dataframe column using case-insensitive matching.
    """

    columns = {
        re.sub(r"[^A-Z0-9]", "", str(col).upper()): col
        for col in df.columns
    }

    normalized_candidates = [
        re.sub(r"[^A-Z0-9]", "", str(name).upper())
        for name in possible_names
    ]

    # Exact match first
    for candidate in normalized_candidates:
        if candidate in columns:
            return columns[candidate]

    # Partial match second
    for candidate in normalized_candidates:
        for normalized_col, original_col in columns.items():
            if candidate and (
                candidate in normalized_col
                or normalized_col in candidate
            ):
                return original_col

    return None


def make_session():
    """
    Create a requests session with browser-like headers.
    """

    session = requests.Session()
    session.headers.update(HEADERS)

    return session


def request_with_retry(session, url, timeout=REQUEST_TIMEOUT):
    """
    Download a URL with retry handling.
    """

    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):

        try:
            response = session.get(
                url,
                timeout=timeout
            )

            response.raise_for_status()

            return response

        except requests.RequestException as exc:

            last_error = exc

            print(
                f"Request attempt {attempt}/{MAX_RETRIES} "
                f"failed: {exc}"
            )

            if attempt < MAX_RETRIES:
                time.sleep(2 * attempt)

    raise RuntimeError(
        f"Unable to download URL after {MAX_RETRIES} attempts:\n"
        f"{url}\n"
        f"Last error: {last_error}"
    )


# ============================================================
# 1. NSE TOP 100 EXTRACTION
# ============================================================

def get_nse100():

    print("\n" + "=" * 65)
    print("DOWNLOADING NSE TOP 100 CONSTITUENTS")
    print("=" * 65)

    session = make_session()

    # Visit the NSE homepage first to establish session cookies.
    try:
        session.get(
            NSE_HOME_URL,
            timeout=REQUEST_TIMEOUT
        )
    except requests.RequestException:
        pass

    response = request_with_retry(
        session,
        NSE_TOP100_URL
    )

    df = pd.read_csv(
        StringIO(response.text),
        dtype=str
    )

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    print("\nNSE CSV columns:")
    print(df.columns.tolist())

    name_col = find_column(
        df,
        ["Company Name", "Company"]
    )

    symbol_col = find_column(
        df,
        ["Symbol", "Trading Symbol"]
    )

    isin_col = find_column(
        df,
        ["ISIN Code", "ISIN"]
    )

    series_col = find_column(
        df,
        ["Series"]
    )

    if not name_col or not symbol_col:
        raise RuntimeError(
            "Unable to identify required NSE Top 100 columns.\n"
            f"Available columns: {df.columns.tolist()}"
        )

    result = pd.DataFrame()

    result["Company Name"] = (
        df[name_col].apply(clean_text)
    )

    result["Symbol"] = (
        df[symbol_col].apply(clean_text)
    )

    if isin_col:
        result["ISIN"] = (
            df[isin_col].apply(clean_isin)
        )
    else:
        result["ISIN"] = ""

    if series_col:
        result["Series"] = (
            df[series_col].apply(clean_text)
        )

        # Keep EQ series if the CSV provides a series column.
        result = result[
            result["Series"].str.upper().eq("EQ")
        ].copy()

    else:
        result["Series"] = ""

    result["BSE Security Code"] = ""
    result["Exchange"] = "NSE"
    result["Index"] = "NSE 100"
    result["Source"] = "Nifty Indices"

    result["_KEY"] = (
        result["Company Name"].apply(
            normalize_company_name
        )
    )

    result = result[
        result["Company Name"].ne("")
        & result["Symbol"].ne("")
        & result["_KEY"].ne("")
    ].copy()

    result = result.drop_duplicates(
        subset="_KEY",
        keep="first"
    )

    if len(result) != EXPECTED_NSE_COUNT:
        raise RuntimeError(
            "NSE Top 100 validation failed.\n"
            f"Expected: {EXPECTED_NSE_COUNT}\n"
            f"Extracted unique companies: {len(result)}\n"
            "Check the downloaded NSE CSV and its columns."
        )

    print("\nNSE extraction completed.")
    print("NSE companies:", len(result))
    print("Unique NSE symbols:", result["Symbol"].nunique())

    print("\nFirst 10 NSE companies:")

    print(
        result[
            ["Company Name", "Symbol", "ISIN"]
        ].head(10).to_string(index=False)
    )

    return result.reset_index(drop=True)


# ============================================================
# 2. BSE 500 EXTRACTION THROUGH DIRECT API
# ============================================================

def get_bse500():

    print("\n" + "=" * 65)
    print("DOWNLOADING BSE 500 CONSTITUENTS")
    print("=" * 65)

    DEBUG_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    session = make_session()

    session.headers.update({
        "Referer": "https://www.bseindices.com/",
        "Accept": "text/csv,*/*",
    })

    response = request_with_retry(
        session,
        BSE_TOP500_API_URL
    )

    print("HTTP Status:", response.status_code)

    print(
        "Content Type:",
        response.headers.get("Content-Type")
    )

    # Save original API response for debugging.
    raw_file = DEBUG_DIR / "bse_api_response.csv"

    raw_file.write_bytes(response.content)

    print("Raw BSE response saved:", raw_file)

    df = pd.read_csv(
        BytesIO(response.content),
        dtype=str,
        encoding="utf-8-sig"
    )

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    print("\nBSE API columns:")
    print(df.columns.tolist())

    name_col = find_column(
        df,
        ["Constituents"]
    )

    code_col = find_column(
        df,
        ["Symbol"]
    )

    sector_col = find_column(
        df,
        ["Macro-Economic Sector"]
    )

    if not name_col or not code_col or not sector_col:

        raise RuntimeError(
            "Unexpected BSE API columns.\n"
            f"Found: {df.columns.tolist()}\n"
            f"Check the raw response: {raw_file}"
        )

    result = pd.DataFrame()

    result["Company Name"] = (
        df[name_col].apply(clean_text)
    )

    result["BSE Security Code"] = (
        df[code_col].apply(clean_text)
    )

    # Retain compatibility with the existing Symbol column.
    result["Symbol"] = result["BSE Security Code"]

    # This BSE API response does not provide ISIN.
    result["ISIN"] = ""

    result["Macro-Economic Sector"] = (
        df[sector_col].apply(clean_text)
    )

    result["Exchange"] = "BSE"
    result["Index"] = "BSE 500"
    result["Source"] = "BSE Indices API"

    result["_KEY"] = (
        result["Company Name"].apply(
            normalize_company_name
        )
    )

    result = result[
        result["Company Name"].ne("")
        & result["BSE Security Code"].ne("")
        & result["_KEY"].ne("")
    ].copy()

    result = result.drop_duplicates(
        subset="_KEY",
        keep="first"
    )

    # Validate BSE security codes.
    duplicate_codes = result[
        result["BSE Security Code"].duplicated(
            keep=False
        )
    ]

    if not duplicate_codes.empty:

        raise RuntimeError(
            "Duplicate BSE security codes detected:\n"
            + duplicate_codes[
                ["Company Name", "BSE Security Code"]
            ].to_string(index=False)
        )

    if len(result) < 450:

        raise RuntimeError(
            f"Only {len(result)} BSE companies extracted. "
            "Expected approximately 500.\n"
            f"Check: {raw_file}"
        )

    print("\nBSE extraction completed.")

    print("BSE company records:", len(result))

    print(
        "Unique BSE security codes:",
        result["BSE Security Code"].nunique()
    )

    print("\nFirst 10 BSE companies:")

    print(
        result[
            [
                "Company Name",
                "BSE Security Code",
                "Macro-Economic Sector",
            ]
        ].head(10).to_string(index=False)
    )

    return result.reset_index(drop=True)


# ============================================================
# 3. NSE EQUITY MASTER EXTRACTION
# ============================================================

def get_nse_equity_master():

    print("\n" + "=" * 65)
    print("DOWNLOADING NSE EQUITY MASTER")
    print("=" * 65)

    session = make_session()

    try:
        session.get(
            NSE_HOME_URL,
            timeout=REQUEST_TIMEOUT
        )
    except requests.RequestException:
        pass

    response = request_with_retry(
        session,
        NSE_EQUITY_MASTER_URL
    )

    df = pd.read_csv(
        StringIO(response.text),
        dtype=str
    )

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    print("\nNSE equity master columns:")
    print(df.columns.tolist())

    name_col = find_column(
        df,
        ["NAME OF COMPANY", "Company Name"]
    )

    symbol_col = find_column(
        df,
        ["SYMBOL", "Trading Symbol"]
    )

    isin_col = find_column(
        df,
        ["ISIN NUMBER", "ISIN"]
    )

    series_col = find_column(
        df,
        ["SERIES"]
    )

    if not name_col or not symbol_col:
        raise RuntimeError(
            "Unable to identify NSE equity-master columns.\n"
            f"Available columns: {df.columns.tolist()}"
        )

    result = pd.DataFrame()

    result["Company Name"] = (
        df[name_col].apply(clean_text)
    )

    result["Symbol"] = (
        df[symbol_col].apply(clean_text)
    )

    if isin_col:
        result["ISIN"] = (
            df[isin_col].apply(clean_isin)
        )
    else:
        result["ISIN"] = ""

    if series_col:

        result["Series"] = (
            df[series_col].apply(clean_text)
        )

        result = result[
            result["Series"].str.upper().eq("EQ")
        ].copy()

    else:
        result["Series"] = ""

    result["BSE Security Code"] = ""
    result["Exchange"] = "NSE"
    result["Index"] = "Additional NSE Listed"
    result["Source"] = "NSE Equity Master"

    result["_KEY"] = (
        result["Company Name"].apply(
            normalize_company_name
        )
    )

    result = result[
        result["Company Name"].ne("")
        & result["Symbol"].ne("")
        & result["_KEY"].ne("")
    ].copy()


    with_isin = result[
        result["ISIN"].ne("")
    ].copy()

    without_isin = result[
        result["ISIN"].eq("")
    ].copy()

    with_isin = with_isin.drop_duplicates(
        subset="ISIN",
        keep="first"
    )

    without_isin = without_isin.drop_duplicates(
        subset="_KEY",
        keep="first"
    )

    result = pd.concat(
        [with_isin, without_isin],
        ignore_index=True
    )

    # Final normalized-name deduplication.
    result = result.drop_duplicates(
        subset="_KEY",
        keep="first"
    )

    result = result.reset_index(drop=True)

    print("\nNSE equity master extraction completed.")

    print("Equity master companies:", len(result))

    print(
        "Companies with ISIN:",
        result["ISIN"].ne("").sum()
    )

    print(
        "Companies without ISIN:",
        result["ISIN"].eq("").sum()
    )

    return result


# ============================================================
# 3. BUILD NSE + BSE INDEX UNIVERSE
# ============================================================

def build_index_universe(nse_df, bse_df):

    print("\n" + "=" * 65)
    print("BUILDING COMBINED NSE + BSE INDEX UNIVERSE")
    print("=" * 65)

    combined = pd.concat(
        [nse_df, bse_df],
        ignore_index=True,
        sort=False
    )

    combined["_KEY"] = (
        combined["_KEY"].fillna("")
    )

    records = []

    for key, group in combined.groupby(
        "_KEY",
        sort=False
    ):

        if not key:
            continue

        first = group.iloc[0].copy()

        exchanges = []
        indices = []
        sources = []

        for value in group["Exchange"].fillna(""):

            for item in str(value).split(" / "):

                item = item.strip()

                if item and item not in exchanges:
                    exchanges.append(item)

        for value in group["Index"].fillna(""):

            for item in str(value).split(" / "):

                item = item.strip()

                if item and item not in indices:
                    indices.append(item)

        for value in group["Source"].fillna(""):

            for item in str(value).split(" / "):

                item = item.strip()

                if item and item not in sources:
                    sources.append(item)

        first["Exchange"] = " / ".join(exchanges)
        first["Index"] = " / ".join(indices)
        first["Source"] = " / ".join(sources)

        # Preserve an available ISIN.
        isins = [
            clean_isin(value)
            for value in group["ISIN"]
        ]

        isins = [
            value for value in isins
            if value
        ]

        first["ISIN"] = (
            isins[0] if isins else ""
        )

        # ----------------------------------------------------
        # IMPORTANT FIX:
        # Preserve BSE security code even if the first row
        # of this group is the NSE record.
        # ----------------------------------------------------

        if "BSE Security Code" in group.columns:

            bse_codes = (
                group["BSE Security Code"]
                .fillna("")
                .astype(str)
                .str.strip()
            )

            bse_codes = [
                code for code in bse_codes
                if code and code.lower() != "nan"
            ]

            first["BSE Security Code"] = (
                bse_codes[0] if bse_codes else ""
            )

        records.append(first)

    index_df = pd.DataFrame(records)

    if index_df.empty:
        raise RuntimeError(
            "Combined NSE/BSE index universe is empty."
        )

    index_df = index_df.reset_index(drop=True)

    # Verify normalized company-name uniqueness.
    duplicate_names = index_df[
        index_df["_KEY"].duplicated(keep=False)
    ]

    if not duplicate_names.empty:
        raise RuntimeError(
            "Duplicate company names found in combined index universe."
        )

    print("\nIndex universe completed.")

    print("NSE records:", len(nse_df))
    print("BSE records:", len(bse_df))

    print(
        "Combined unique index companies:",
        len(index_df)
    )

    print(
        "Companies appearing in both exchanges:",
        len(nse_df) + len(bse_df) - len(index_df)
    )

    return index_df


# ============================================================
# 4. FIND ADDITIONAL COMPANIES
# ============================================================

def get_additional_companies(index_df, master_df):

    print("\n" + "=" * 65)
    print("IDENTIFYING ADDITIONAL NSE COMPANIES")
    print("=" * 65)

    required = TARGET_TOTAL - len(index_df)

    print("Target total:", TARGET_TOTAL)
    print("Existing index companies:", len(index_df))
    print("Additional companies required:", required)

    if required < 0:

        raise RuntimeError(
            f"Index universe already contains {len(index_df)} "
            f"unique companies, exceeding target {TARGET_TOTAL}."
        )

    if required == 0:

        print("No additional companies required.")

        return master_df.head(0).copy()

    existing_names = set(
        index_df["_KEY"].dropna().astype(str)
    )

    existing_isins = set(
        index_df["ISIN"]
        .apply(clean_isin)
    )

    existing_isins.discard("")

    candidates = master_df.copy()

    candidates = candidates[
        ~candidates["_KEY"].isin(existing_names)
    ].copy()

    # Exclude existing ISINs.
    candidates = candidates[
        ~candidates["ISIN"].apply(clean_isin).isin(
            existing_isins
        )
    ].copy()

    # Remove duplicate company names.
    candidates = candidates.drop_duplicates(
        subset="_KEY",
        keep="first"
    )

    # Build a unique selection.
    selected_rows = []

    selected_names = set(existing_names)
    selected_isins = set(existing_isins)

    for _, row in candidates.iterrows():

        key = clean_text(row["_KEY"])
        isin = clean_isin(row["ISIN"])

        if not key:
            continue

        if key in selected_names:
            continue

        if isin and isin in selected_isins:
            continue

        selected_rows.append(row)

        selected_names.add(key)

        if isin:
            selected_isins.add(isin)

        if len(selected_rows) == required:
            break

    if len(selected_rows) < required:

        raise RuntimeError(
            "\nUnable to reach the target of 600 unique companies.\n"
            f"Required additional companies: {required}\n"
            f"Available valid additional companies: "
            f"{len(selected_rows)}\n"
            f"Final possible total: "
            f"{len(index_df) + len(selected_rows)}\n\n"
            "The script will not fabricate company records. "
            "Check the NSE equity master or its filters."
        )

    additional_df = pd.DataFrame(
        selected_rows
    ).reset_index(drop=True)

    additional_df["Exchange"] = "NSE"
    additional_df["Index"] = "Additional NSE Listed"
    additional_df["Source"] = "NSE Equity Master"

    print("\nAdditional companies selected:")
    print(len(additional_df))

    print("\nFirst 10 additional companies:")

    print(
        additional_df[
            ["Company Name", "Symbol", "ISIN"]
        ].head(10).to_string(index=False)
    )

    return additional_df


# ============================================================
# 5. BUILD FINAL 600 COMPANIES
# ============================================================

def build_final_list(index_df, additional_df):

    print("\n" + "=" * 65)
    print("BUILDING FINAL COMPANY LIST")
    print("=" * 65)

    final_df = pd.concat(
        [index_df, additional_df],
        ignore_index=True,
        sort=False
    )

    final_df["_KEY"] = (
        final_df["_KEY"].fillna("")
    )

    final_df["ISIN"] = (
        final_df["ISIN"].apply(clean_isin)
    )

    # Validate company-name uniqueness.
    duplicate_names = final_df[
        final_df["_KEY"].duplicated(keep=False)
    ]

    if not duplicate_names.empty:

        raise RuntimeError(
            "Duplicate normalized company names found "
            "in the final list:\n"
            + duplicate_names[
                ["Company Name", "Exchange", "ISIN"]
            ].to_string(index=False)
        )

    # Validate nonblank ISIN uniqueness.
    nonblank_isins = final_df[
        final_df["ISIN"].ne("")
    ]

    duplicate_isins = nonblank_isins[
        nonblank_isins["ISIN"].duplicated(keep=False)
    ]

    if not duplicate_isins.empty:

        raise RuntimeError(
            "Duplicate ISIN values found in the final list:\n"
            + duplicate_isins[
                ["Company Name", "ISIN"]
            ].to_string(index=False)
        )

    # Exact target validation.
    if len(final_df) != TARGET_TOTAL:

        raise RuntimeError(
            f"Final company count is {len(final_df)}. "
            f"Expected exactly {TARGET_TOTAL}."
        )

    if final_df["_KEY"].nunique() != TARGET_TOTAL:

        raise RuntimeError(
            "Final normalized company-name count "
            "does not equal 600."
        )

    print("\nFinal validation successful.")

    print("Total company records:", len(final_df))

    print(
        "Unique company names:",
        final_df["_KEY"].nunique()
    )

    print(
        "Unique nonblank ISINs:",
        final_df.loc[
            final_df["ISIN"].ne(""),
            "ISIN"
        ].nunique()
    )

    return final_df.reset_index(drop=True)


# ============================================================
# 6. BUILD OVERLAP SHEET
# ============================================================

def build_overlap(nse_df, bse_df):

    nse_keys = set(
        nse_df["_KEY"].dropna()
    )

    bse_keys = set(
        bse_df["_KEY"].dropna()
    )

    overlap_keys = nse_keys.intersection(
        bse_keys
    )

    nse_overlap = nse_df[
        nse_df["_KEY"].isin(overlap_keys)
    ].copy()

    bse_overlap = bse_df[
        bse_df["_KEY"].isin(overlap_keys)
    ].copy()

    nse_overlap = nse_overlap.rename(
        columns={
            "Company Name": "NSE Company Name",
            "Symbol": "NSE Symbol",
            "ISIN": "NSE ISIN",
        }
    )

    bse_overlap = bse_overlap.rename(
        columns={
            "Company Name": "BSE Company Name",
            "BSE Security Code": "BSE Security Code",
            "Symbol": "BSE Symbol",
            "Macro-Economic Sector": "BSE Sector",
        }
    )

    overlap = nse_overlap[
        [
            "_KEY",
            "NSE Company Name",
            "NSE Symbol",
            "NSE ISIN",
        ]
    ].merge(
        bse_overlap[
            [
                "_KEY",
                "BSE Company Name",
                "BSE Symbol",
                "BSE Security Code",
                "BSE Sector",
            ]
        ],
        on="_KEY",
        how="inner"
    )

    overlap = overlap.drop(
        columns=["_KEY"]
    )

    overlap = overlap.reset_index(drop=True)

    return overlap


# ============================================================
# 7. BUILD SUMMARY SHEET
# ============================================================

def build_summary(
    nse_df,
    bse_df,
    index_df,
    overlap_df,
    additional_df,
    final_df
):

    summary_data = [
        {
            "Metric": "Target Company Count",
            "Value": TARGET_TOTAL,
        },
        {
            "Metric": "NSE Top 100 Records",
            "Value": len(nse_df),
        },
        {
            "Metric": "BSE 500 Records",
            "Value": len(bse_df),
        },
        {
            "Metric": "NSE-BSE Overlap",
            "Value": len(overlap_df),
        },
        {
            "Metric": "Unique NSE-BSE Index Companies",
            "Value": len(index_df),
        },
        {
            "Metric": "Additional NSE Companies",
            "Value": len(additional_df),
        },
        {
            "Metric": "Final Company Count",
            "Value": len(final_df),
        },
        {
            "Metric": "Final Unique Company Names",
            "Value": final_df["_KEY"].nunique(),
        },
        {
            "Metric": "Companies with ISIN",
            "Value": final_df["ISIN"].ne("").sum(),
        },
        {
            "Metric": "Companies without ISIN",
            "Value": final_df["ISIN"].eq("").sum(),
        },
        {
            "Metric": "Companies with BSE Security Code",
            "Value": final_df[
                "BSE Security Code"
            ].fillna("").astype(str).str.strip().ne("").sum(),
        },
        {
            "Metric": "Validation Status",
            "Value": (
                "PASSED"
                if len(final_df) == TARGET_TOTAL
                else "FAILED"
            ),
        },
    ]

    return pd.DataFrame(summary_data)


# ============================================================
# 8. CLEAN OUTPUT COLUMNS
# ============================================================

def prepare_output(df):

    result = df.copy()

    # Ensure all expected columns exist.
    expected_columns = [
        "Company Name",
        "Symbol",
        "ISIN",
        "Exchange",
        "Index",
        "Source",
        "BSE Security Code",
        "Series",
        "Macro-Economic Sector",
    ]

    for col in expected_columns:

        if col not in result.columns:
            result[col] = ""

    # Remove internal matching key.
    if "_KEY" in result.columns:
        result = result.drop(columns=["_KEY"])

    # Remove duplicate output columns if any.
    result = result.loc[
        :,
        ~result.columns.duplicated()
    ]

    # Keep standard columns first.
    other_columns = [
        col for col in result.columns
        if col not in expected_columns
    ]

    result = result[
        expected_columns + other_columns
    ]

    return result


# ============================================================
# 9. SAVE FINAL EXCEL
# ============================================================

def save_excel(
    nse_df,
    bse_df,
    overlap_df,
    index_df,
    additional_df,
    final_df,
    summary_df
):

    print("\n" + "=" * 65)
    print("SAVING FINAL EXCEL FILE")
    print("=" * 65)

    BASE_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    nse_output = prepare_output(nse_df)
    bse_output = prepare_output(bse_df)
    index_output = prepare_output(index_df)
    additional_output = prepare_output(additional_df)
    final_output = prepare_output(final_df)

    overlap_output = overlap_df.copy()

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        final_output.to_excel(
            writer,
            sheet_name="All Companies",
            index=False
        )

        nse_output.to_excel(
            writer,
            sheet_name="NSE 100",
            index=False
        )

        bse_output.to_excel(
            writer,
            sheet_name="BSE 500",
            index=False
        )

        overlap_output.to_excel(
            writer,
            sheet_name="Overlap",
            index=False
        )

        additional_output.to_excel(
            writer,
            sheet_name="Additional Companies",
            index=False
        )

        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        # Format worksheets.
        for worksheet in writer.sheets.values():

            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = (
                worksheet.dimensions
            )

            for column_cells in worksheet.columns:

                max_length = 0

                column_letter = (
                    column_cells[0].column_letter
                )

                for cell in column_cells:

                    value = cell.value

                    if value is not None:
                        max_length = max(
                            max_length,
                            len(str(value))
                        )

                worksheet.column_dimensions[
                    column_letter
                ].width = min(max_length + 3, 45)

    print("\nExcel file saved successfully.")

    print("Output:", OUTPUT_FILE)

    print("\nSheets created:")

    print("1. All Companies")
    print("2. NSE 100")
    print("3. BSE 500")
    print("4. Overlap")
    print("5. Additional Companies")
    print("6. Summary")


# ============================================================
# 10. MAIN EXECUTION
# ============================================================

def main():

    start_time = time.time()

    print("\n")
    print("=" * 75)
    print("NSE + BSE TOP 600 COMPANY EXTRACTION")
    print("=" * 75)

    print("Project directory:", BASE_DIR)
    print("Target companies:", TARGET_TOTAL)
    print("Output file:", OUTPUT_FILE)

    try:

        # ----------------------------------------------------
        # STEP 1: NSE Top 100
        # ----------------------------------------------------

        nse_df = get_nse100()

        # ----------------------------------------------------
        # STEP 2: BSE 500
        # ----------------------------------------------------

        bse_df = get_bse500()

        # ----------------------------------------------------
        # STEP 3: Combined index universe
        # ----------------------------------------------------

        index_df = build_index_universe(
            nse_df,
            bse_df
        )

        # ----------------------------------------------------
        # STEP 4: NSE equity master
        # ----------------------------------------------------

        master_df = get_nse_equity_master()

        # ----------------------------------------------------
        # STEP 5: Additional companies
        # ----------------------------------------------------

        additional_df = get_additional_companies(
            index_df,
            master_df
        )

        # ----------------------------------------------------
        # STEP 6: Final 600 validation
        # ----------------------------------------------------

        final_df = build_final_list(
            index_df,
            additional_df
        )

        # ----------------------------------------------------
        # STEP 7: Overlap sheet
        # ----------------------------------------------------

        overlap_df = build_overlap(
            nse_df,
            bse_df
        )

        # ----------------------------------------------------
        # STEP 8: Summary
        # ----------------------------------------------------

        summary_df = build_summary(
            nse_df,
            bse_df,
            index_df,
            overlap_df,
            additional_df,
            final_df
        )

        # ----------------------------------------------------
        # STEP 9: Save Excel
        # ----------------------------------------------------

        save_excel(
            nse_df,
            bse_df,
            overlap_df,
            index_df,
            additional_df,
            final_df,
            summary_df
        )

        elapsed_time = time.time() - start_time

        print("\n" + "=" * 75)
        print("PROCESS COMPLETED SUCCESSFULLY")
        print("=" * 75)

        print("NSE companies:", len(nse_df))
        print("BSE companies:", len(bse_df))
        print("Overlap:", len(overlap_df))
        print("Unique index companies:", len(index_df))
        print("Additional companies:", len(additional_df))
        print("Final unique companies:", len(final_df))

        print(
            "Execution time:",
            round(elapsed_time, 2),
            "seconds"
        )

        print("\nFinal Excel:")
        print(OUTPUT_FILE)

    except Exception as exc:

        print("\n" + "=" * 75)
        print("PROCESS FAILED")
        print("=" * 75)

        print("\nError:")
        print(str(exc))

        print(
            "\nThe script stopped without creating "
            "a misleading 600-company output."
        )

        raise


if __name__ == "__main__":
    main()
