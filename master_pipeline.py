import os
import sys
import time
import subprocess
from pathlib import Path


# ============================================================
# 1. CENTRAL CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
os.environ["PROJECT_DIR"] = str(BASE_DIR)

PYTHON = sys.executable


# ------------------------------------------------------------
# COMPANY MASTER
# ------------------------------------------------------------

COMPANY_FILE = (
    BASE_DIR / "NSE_BSE_Top_600s_Companies.xlsx"
)

COMPANY_SHEET = "All Companies"
# Complete company master contains 600 companies.
EXPECTED_COMPANIES = 600


# ------------------------------------------------------------
# FINANCIAL EXTRACTION
# ------------------------------------------------------------

FINANCIAL_FILE = (
    BASE_DIR
    / "output"
    / "companies_final_financial_sector_ranked.xlsx"
)


# ------------------------------------------------------------
# NEWS EXTRACTION
# ------------------------------------------------------------

NEWS_FILE = (
    BASE_DIR / "output" / "news_analysis.xlsx"
)


# ------------------------------------------------------------
# ANNUAL REPORT URLS
# ------------------------------------------------------------

URL_FILE = (
    BASE_DIR / "output" / "Companies_urls.xlsx"
)

URL_COLUMN = "Alternative pdf url"


# ------------------------------------------------------------
# ANNUAL REPORT PDF FOLDER
# ------------------------------------------------------------

PDF_DIR = (
    BASE_DIR / "annual_reports"
)


# ------------------------------------------------------------
# TECHNOLOGY EXTRACTION
# ------------------------------------------------------------

TECH_NOTEBOOK = (
    BASE_DIR / "annual_report_rag_8GB.ipynb"
)

TECH_DIR = (
    BASE_DIR / "tech_matrix_output"
)

TECH_FILE = (
    TECH_DIR / "NSE_BSE Companies.xlsx"
)
TEST_COMPANIES = 600
TECH_SHEET = "Tech Matrix"

EXPECTED_TECH_PARAMETERS = 15

EXPECTED_TECH_COMPANIES = 600


# ------------------------------------------------------------
# STREAMLIT DASHBOARD
# ------------------------------------------------------------

APP_FILE = (
    BASE_DIR / "app.py"
)


# ------------------------------------------------------------
# PIPELINE CONTROL
# ------------------------------------------------------------

RUN_COMPANY_GENERATION = True

RUN_FINANCIAL = True

RUN_NEWS = True

RUN_URL_EXTRACTION = True

RUN_PDF_DOWNLOAD = True

RUN_TECHNOLOGY = True

RUN_DASHBOARD = True


# ============================================================
# 2. CREATE REQUIRED DIRECTORIES
# ============================================================

def create_directories():

    directories = [
        BASE_DIR / "output",
        BASE_DIR / "annual_reports",
        BASE_DIR / "tech_matrix_output",
        BASE_DIR / "output" / "pipeline_notebooks",
    ]

    for directory in directories:

        directory.mkdir(
            parents=True,
            exist_ok=True
        )

    print("[OK] Required directories created.")


# ============================================================
# 3. COMMON FUNCTIONS
# ============================================================

def print_header(title):

    print("\n")
    print("=" * 80)
    print(title.center(80))
    print("=" * 80)


def run_command(command, stage_name, cwd=BASE_DIR):

    print_header(f"STARTING: {stage_name}")

    print("Command:")
    print(" ".join(str(x) for x in command))

    start = time.time()

    result = subprocess.run(
        command,
        cwd=str(cwd),
        check=False
    )

    elapsed = round(
        time.time() - start,
        2
    )

    if result.returncode != 0:

        raise RuntimeError(
            f"{stage_name} failed with exit code "
            f"{result.returncode}"
        )

    print(f"\n[COMPLETED] {stage_name}")
    print(f"Time taken: {elapsed} seconds")


def run_python_script(script_name):

    script_path = (
        BASE_DIR / script_name
    )

    if not script_path.exists():

        raise FileNotFoundError(
            f"Script not found: {script_path}"
        )

    run_command(
        [
            PYTHON,
            "-u",
            str(script_path)
        ],
        script_name
    )


def require_file(path, stage_name):

    if not path.exists():

        raise FileNotFoundError(
            f"\n{stage_name} did not generate its expected file:\n"
            f"{path}"
        )

    if not path.is_file():

        raise RuntimeError(
            f"{stage_name}: Expected a file but found:\n{path}"
        )

    if path.stat().st_size == 0:

        raise RuntimeError(
            f"{stage_name} generated an empty file:\n{path}"
        )

    print(f"[VALIDATED] {path}")


def load_excel(path, sheet_name=0):

    import pandas as pd

    require_file(
        path,
        "Excel input"
    )

    return pd.read_excel(
        path,
        sheet_name=sheet_name,
        dtype=str
    ).fillna("")


# ============================================================
# 4. VALIDATE 600-COMPANY MASTER EXCEL
# ============================================================

def validate_company_file():
    print("DEBUG - EXPECTED_COMPANIES:", EXPECTED_COMPANIES)
    print("DEBUG - COMPANY_FILE:", COMPANY_FILE)
    print_header("VALIDATING 600-COMPANY MASTER")

    df = load_excel(
        COMPANY_FILE,
        sheet_name=COMPANY_SHEET
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    if "Company Name" not in df.columns:

        raise ValueError(
            "All Companies sheet does not contain "
            "'Company Name' column."
        )

    names = (
        df["Company Name"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    names = names[
        names.ne("")
    ]

    unique_names = names.nunique()

    print("\nCompany master validation:")
    print("Total rows:", len(df))
    print("Non-empty company names:", len(names))
    print("Unique company names:", unique_names)

    if len(df) != EXPECTED_COMPANIES:

        raise RuntimeError(
            f"Expected {EXPECTED_COMPANIES} company rows, "
            f"but found {len(df)}."
        )

    if unique_names != EXPECTED_COMPANIES:

        raise RuntimeError(
            f"Expected {EXPECTED_COMPANIES} unique companies, "
            f"but found {unique_names}."
        )

    print(
        f"\n[PASSED] Exactly {EXPECTED_COMPANIES} "
        "unique companies found."
    )

    return df


# ============================================================
# 5. FINANCIAL DATA EXTRACTION
# ============================================================

def run_financial_extraction():

    print_header(
        "FINANCIAL DATA EXTRACTION"
    )

    if not RUN_FINANCIAL:

        print("[SKIPPED] Financial extraction disabled.")

        require_file(
            FINANCIAL_FILE,
            "Existing financial workbook"
        )

        return

    run_python_script(
        "financial.py"
    )

    require_file(
        FINANCIAL_FILE,
        "Financial extraction"
    )

    df = load_excel(
        FINANCIAL_FILE
    )

    print("\nFinancial extraction summary:")
    print("Financial records:", len(df))

    company_col = next(
        (
            col for col in [
                "Company_Name",
                "Company Name",
                "Company"
            ]
            if col in df.columns
        ),
        None
    )

    if company_col:

        unique_count = (
            df[company_col]
            .astype(str)
            .str.strip()
            .replace("", None)
            .nunique()
        )

        print(
            "Unique financial companies:",
            unique_count
        )

    else:

        print(
            "[WARNING] Company name column was not detected "
            "in the financial workbook."
        )


# ============================================================
# 6. COMPANY NEWS EXTRACTION
# ============================================================

def run_news_extraction():

    print_header(
        "COMPANY NEWS EXTRACTION"
    )

    if not RUN_NEWS:

        print("[SKIPPED] News extraction disabled.")

        require_file(
            NEWS_FILE,
            "Existing news workbook"
        )

        return

    run_python_script(
        "news_intelligence.py"
    )

    require_file(
        NEWS_FILE,
        "News extraction"
    )

    df = load_excel(
        NEWS_FILE
    )

    print("\nNews extraction summary:")
    print("News records:", len(df))

    company_col = next(
        (
            col for col in [
                "Company",
                "Company Name",
                "Company_Name"
            ]
            if col in df.columns
        ),
        None
    )

    if company_col:

        unique_count = (
            df[company_col]
            .astype(str)
            .str.strip()
            .replace("", None)
            .nunique()
        )

        print(
            "Unique companies with news:",
            unique_count
        )


# ============================================================
# 7. ANNUAL REPORT URL EXTRACTION
# ============================================================

def run_url_extraction():

    print_header(
        "ANNUAL REPORT URL EXTRACTION"
    )

    if not RUN_URL_EXTRACTION:

        print("[SKIPPED] URL extraction disabled.")

        require_file(
            URL_FILE,
            "Existing annual-report URL file"
        )

        return

    run_python_script(
        "save_urls.py"
    )

    require_file(
        URL_FILE,
        "Annual-report URL extraction"
    )

    df = load_excel(
        URL_FILE
    )

    print("\nAnnual-report URL summary:")
    print("Companies in URL workbook:", len(df))

    if URL_COLUMN not in df.columns:

        raise ValueError(
            f"URL workbook is missing the "
            f"'{URL_COLUMN}' column."
        )

    valid_urls = (
        df[URL_COLUMN]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.startswith("http")
    )

    print(
        "Companies with PDF URLs:",
        int(valid_urls.sum())
    )

    print(
        "Companies without PDF URLs:",
        int((~valid_urls).sum())
    )

    if valid_urls.sum() == 0:

        raise RuntimeError(
            "No valid annual-report PDF URLs were found."
        )


# ============================================================
# 8. DOWNLOAD ANNUAL REPORTS
# ============================================================

def run_pdf_downloader():

    print_header(
        "ANNUAL REPORT PDF DOWNLOADING"
    )

    if not RUN_PDF_DOWNLOAD:

        print("[SKIPPED] PDF downloading disabled.")

    else:

        run_python_script(
            "download_urls.py"
        )

    if not PDF_DIR.is_dir():

        raise FileNotFoundError(
            f"Annual reports folder not found: {PDF_DIR}"
        )

    pdf_files = list(
        PDF_DIR.rglob("*.pdf")
    )

    print("\nPDF download summary:")
    print("PDF folder:", PDF_DIR)
    print("PDF files available:", len(pdf_files))

    if not pdf_files:

        raise RuntimeError(
            "No annual-report PDFs were found. "
            "Technology extraction cannot continue."
        )

    print(
        "[PASSED] Annual-report PDFs are available."
    )


# ============================================================
# 9. TECHNOLOGY EXTRACTION NOTEBOOK
# ============================================================

def run_technology_extraction():

    print_header(
        "TECHNOLOGY EXTRACTION - 15 PARAMETERS"
    )

    if not RUN_TECHNOLOGY:

        print("[SKIPPED] Technology extraction disabled.")

        require_file(
            TECH_FILE,
            "Existing technology workbook"
        )

        return

    require_file(
        TECH_NOTEBOOK,
        "Technology notebook"
    )

    require_file(
        COMPANY_FILE,
        "Technology company input"
    )

    if not PDF_DIR.is_dir():

        raise FileNotFoundError(
            f"Annual reports folder not found: {PDF_DIR}"
        )

    pdf_files = list(
        PDF_DIR.rglob("*.pdf")
    )

    if not pdf_files:

        raise RuntimeError(
            "No PDF reports available for technology extraction."
        )

    # --------------------------------------------------------
    # EXECUTED NOTEBOOK OUTPUT
    # --------------------------------------------------------

    executed_dir = (
        BASE_DIR
        / "output"
        / "pipeline_notebooks"
    )

    executed_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    executed_notebook = (
        executed_dir
        / "annual_report_rag_8GB_executed.ipynb"
    )

    # --------------------------------------------------------
    # NOTEBOOK EXECUTION
    # --------------------------------------------------------

    command = [
        PYTHON,
        "-m",
        "nbconvert",
        "--to",
        "notebook",
        "--execute",
        str(TECH_NOTEBOOK),
        "--output",
        executed_notebook.name,
        "--output-dir",
        str(executed_dir),
        "--ExecutePreprocessor.timeout=-1",
        "--ExecutePreprocessor.kernel_name=python3",
    ]

    run_command(
        command,
        "Technology extraction notebook"
    )

    # --------------------------------------------------------
    # VALIDATE EXECUTED NOTEBOOK
    # --------------------------------------------------------

    require_file(
        executed_notebook,
        "Executed technology notebook"
    )

    # --------------------------------------------------------
    # VALIDATE TECHNOLOGY EXCEL
    # --------------------------------------------------------

    require_file(
        TECH_FILE,
        "Technology extraction"
    )

    import pandas as pd

    matrix = pd.read_excel(
        TECH_FILE,
        sheet_name=TECH_SHEET,
        dtype=str
    ).fillna("")

    matrix.columns = (
        matrix.columns
        .astype(str)
        .str.strip()
    )

    print("\nTechnology extraction summary:")
    print("Total technology records:", len(matrix))

    if "Company" not in matrix.columns:

        raise ValueError(
            "Technology workbook is missing the Company column."
        )

    unique_companies = (
        matrix["Company"]
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .nunique()
    )

    print(
        "Unique companies:",
        unique_companies
    )

    # --------------------------------------------------------
    # IDENTIFY TECHNOLOGY PARAMETER COLUMNS
    # --------------------------------------------------------

    metadata_columns = [
        "No.",
        "Company",
        "Sector",
        "PDF File"
    ]

    parameter_columns = [
        col for col in matrix.columns
        if col not in metadata_columns
    ]

    print(
        "Technology parameter columns:",
        len(parameter_columns)
    )

    # --------------------------------------------------------
    # VALIDATE TECHNOLOGY COMPANY COUNT
    # --------------------------------------------------------

    if unique_companies != EXPECTED_TECH_COMPANIES:

        raise RuntimeError(
            "Technology matrix does not contain the expected "
            f"{EXPECTED_TECH_COMPANIES} unique companies. "
            f"Found: {unique_companies}"
        )

    # --------------------------------------------------------
    # VALIDATE 15 PARAMETERS
    # --------------------------------------------------------

    if len(parameter_columns) != EXPECTED_TECH_PARAMETERS:

        raise RuntimeError(
            f"Expected {EXPECTED_TECH_PARAMETERS} technology "
            f"parameters, found {len(parameter_columns)}."
        )

    print(
        "\n[PASSED] Technology matrix contains "
        f"{EXPECTED_TECH_COMPANIES} unique companies "
        f"and {EXPECTED_TECH_PARAMETERS} parameters."
    )


# ============================================================
# 10. FINAL DATASET VALIDATION
# ============================================================

def validate_final_outputs():

    print_header(
        "FINAL DATASET VALIDATION"
    )

    files = {
        "600-company master": COMPANY_FILE,
        "Financial data": FINANCIAL_FILE,
        "News data": NEWS_FILE,
        "Annual-report URLs": URL_FILE,
        "Technology data": TECH_FILE,
    }

    for name, path in files.items():

        require_file(
            path,
            name
        )

    print("\nAll required datasets are available.")

    print("\nFinal output locations:")

    for name, path in files.items():

        print(f"{name}: {path}")

    print(
        "\n[PASSED] Final output file validation completed."
    )


# ============================================================
# 11. LAUNCH STREAMLIT DASHBOARD
# ============================================================

def launch_dashboard():

    print_header(
        "LAUNCHING CORPORATE INTELLIGENCE DASHBOARD"
    )

    require_file(
        APP_FILE,
        "Streamlit dashboard"
    )

    command = [
        PYTHON,
        "-m",
        "streamlit",
        "run",
        str(APP_FILE)
    ]

    print("Starting Streamlit...")
    print("Dashboard file:", APP_FILE)

    print(
        "\nThe dashboard will open in your browser."
    )

    subprocess.Popen(
        command,
        cwd=str(BASE_DIR)
    )

    print(
        "\n[STARTED] Streamlit dashboard process."
    )


# ============================================================
# 12. MAIN PIPELINE
# ============================================================

def main():

    start_time = time.time()

    print_header(
        "CORPORATE INTELLIGENCE - COMPLETE AUTOMATED PIPELINE"
    )

    print("Project directory:", BASE_DIR)
    print("Python executable:", PYTHON)
    print("Target master companies:", EXPECTED_COMPANIES)
    print("Technology test companies:", EXPECTED_TECH_COMPANIES)
    print("Technology parameters:", EXPECTED_TECH_PARAMETERS)

    try:

        # ----------------------------------------------------
        # STEP 0: CREATE DIRECTORIES
        # ----------------------------------------------------

        create_directories()

        # ----------------------------------------------------
        # STEP 1: GENERATE COMPANY MASTER
        # ----------------------------------------------------

        if RUN_COMPANY_GENERATION:

            run_python_script(
                "fetcher.py"
            )

        else:

            print(
                "\n[SKIPPED] Company generation disabled. "
                "Using existing company Excel."
            )

        # ----------------------------------------------------
        # STEP 2: VALIDATE COMPANY MASTER
        # ----------------------------------------------------

        validate_company_file()

        # ----------------------------------------------------
        # STEP 3: FINANCIAL EXTRACTION
        # ----------------------------------------------------

        run_financial_extraction()

        # ----------------------------------------------------
        # STEP 4: NEWS EXTRACTION
        # ----------------------------------------------------

        run_news_extraction()

        # ----------------------------------------------------
        # STEP 5: ANNUAL REPORT URL EXTRACTION
        # ----------------------------------------------------

        run_url_extraction()

        # ----------------------------------------------------
        # STEP 6: DOWNLOAD ANNUAL REPORT PDFs
        # ----------------------------------------------------

        run_pdf_downloader()

        # ----------------------------------------------------
        # STEP 7: TECHNOLOGY EXTRACTION
        # ----------------------------------------------------

        run_technology_extraction()

        # ----------------------------------------------------
        # STEP 8: FINAL DATASET VALIDATION
        # ----------------------------------------------------

        validate_final_outputs()

        # ----------------------------------------------------
        # TOTAL EXECUTION TIME
        # ----------------------------------------------------

        elapsed = round(
            (time.time() - start_time) / 60,
            2
        )

        print_header(
            "ALL PIPELINE STAGES COMPLETED"
        )

        print(
            "Total execution time:",
            elapsed,
            "minutes"
        )

        print(
            "\nAll output datasets are ready."
        )

        # ----------------------------------------------------
        # STEP 9: LAUNCH DASHBOARD
        # ----------------------------------------------------

        if RUN_DASHBOARD:

            launch_dashboard()

        else:

            print(
                "\n[SKIPPED] Dashboard launch disabled."
            )

    except Exception as error:

        print_header(
            "PIPELINE STOPPED - ERROR"
        )

        print("\nReason:")
        print(error)

        print(
            "\nThe pipeline stopped to prevent later stages "
            "from using incomplete data."
        )

        raise


# ============================================================
# 13. RUN PIPELINE
# ============================================================

if __name__ == "__main__":

    main()
