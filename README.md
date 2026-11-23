# Corporate Intelligence – NSE/BSE Top 600 Companies

An end-to-end pipeline and Streamlit dashboard that, for the top ~600 NSE/BSE listed companies:

1. Builds the company master list (NSE 100 + BSE 500)
2. Extracts financial data and sector rankings
3. Finds and downloads annual report PDFs
4. Extracts 15 technology-adoption parameters from the reports using a RAG pipeline (Groq LLM)
5. Serves everything in a dashboard with sector comparison and an AI assistant

---

## Requirements

- Python 3.10 – 3.12
- Internet access (NSE/BSE sites, Serper, Groq)
- Optional but strongly recommended: NVIDIA GPU (~8 GB VRAM) for the notebook's embedding and reranker models. It falls back to CPU, which is much slower.
- API keys:
  - **Groq** (https://console.groq.com) for LLM extraction and the chatbot

---

## 1. Setup

### Create and activate a virtual environment

**Windows (PowerShell)**
```powershell
cd path\to\Corporate__Intelligence
python -m venv .venv
.venv\Scripts\Activate.ps1
```

**macOS / Linux**
```bash
cd path/to/Corporate__Intelligence
python3 -m venv .venv
source .venv/bin/activate
```

### Install dependencies

```bash
pip install --upgrade pip
pip install -r requirements.txt
```


> **GPU users:** install the CUDA build of PyTorch from https://pytorch.org/get-started/locally/ *before* running `pip install -r requirements.txt`.

### Install the Playwright browser (one time)

```bash
playwright install chromium
```

### Register the notebook kernel (one time)

`master_pipeline.py` runs the notebook with the `python3` kernel:

```bash
python -m ipykernel install --user --name python3
```

---

## 2. Configure API keys

Create a file named `.env` in the project root:

```env
SERPER_API_KEY=your_serper_key
GROQ_API_KEY=your_groq_key

# Optional (dashboard chatbot)
GROQ_MODEL=openai/gpt-oss-20b
GROQ_REASONING_EFFORT=low
```

Create `groq_key.txt` in the project root containing **only** your Groq key on one line (the notebook reads this file):

```
your_groq_key
```

**Never commit `.env` or `groq_key.txt`.** Confirm `.gitignore` contains:

```
.env
groq_key.txt
.venv/
__pycache__/
```

---

## 3. Notebook path configuration

Open `annual_report_rag_8GB.ipynb`, find the settings cell ("EDIT ONLY THIS CELL BEFORE EACH RUN"), and replace the hardcoded Windows path:

```python
# Before
BASE_DIR = Path(r"D:\GENXAI\Corporate__Intelligence")

# After
import os
BASE_DIR = Path(os.environ.get("PROJECT_DIR", Path.cwd()))
```

Then add this to `master_pipeline.py` directly after `BASE_DIR = Path(__file__).resolve().parent`:

```python
import os
os.environ["PROJECT_DIR"] = str(BASE_DIR)
```

---

## 4. Run the full pipeline

> **Do a test run first.** The full run covers 600 companies and takes many hours (the LLM step alone pauses 6 s per company). See [Test run](#test-run-recommended-first).

```bash
python master_pipeline.py
```

Stages run in this order:

| # | Stage | Script | Output |
|---|-------|--------|--------|
| 1 | Company list | `fetcher.py` | `NSE_BSE_Top_600s_Companies.xlsx` (sheet `All Companies`) |
| 2 | Financials | `financial.py` | `output/companies_final_financial_sector_ranked.xlsx` |
| 3 | Report URLs | `save_urls.py` | `output/Companies_urls.xlsx` |
| 4 | PDF download | `download_urls.py` | `annual_reports/*.pdf` |
| 5 | Technology matrix | notebook | `tech_matrix_output/NSE_BSE Companies.xlsx` (sheet `Tech Matrix`) |
| 6 | Dashboard | `app.py` | opens in browser |

An executed copy of the notebook is saved to `output/pipeline_notebooks/`.

### Skip stages you have already completed

At the top of `master_pipeline.py`, set any finished stage to `False`:

```python
RUN_COMPANY_GENERATION = True
RUN_FINANCIAL = True
RUN_URL_EXTRACTION = True
RUN_PDF_DOWNLOAD = True
RUN_TECHNOLOGY = True
RUN_DASHBOARD = True
```

If you set `RUN_COMPANY_GENERATION = False`, the company workbook must already have a sheet named **`All Companies`** (the scripts read that sheet name).

### Run stages individually

```bash
python fetcher.py              # company list
python financial.py            # financial data
python save_urls.py            # annual report URLs
python download_urls.py        # download PDFs
```

Technology extraction manually (from the project folder):

```bash
jupyter notebook annual_report_rag_8GB.ipynb
```

or headless:

```bash
python -m nbconvert --to notebook --execute annual_report_rag_8GB.ipynb --output annual_report_rag_8GB_executed.ipynb --output-dir output/pipeline_notebooks --ExecutePreprocessor.timeout=-1
```

---

## Test run (recommended first)

Limit every stage to a few companies before committing to all 600:

| File | Setting | Test value |
|------|---------|-----------|
| `download_urls.py` | `END_ROW` | `5` |
| `master_pipeline.py` | `TEST_COMPANIES`, `EXPECTED_COMPANIES`, `EXPECTED_TECH_COMPANIES` | `5` (so validation does not expect 600) |
| Notebook settings cell | `END_COMPANY` | `5` |

Run `python master_pipeline.py`, confirm the outputs look right, then restore the values to `600`.

The notebook skips companies already in the Tech Matrix. Set `FORCE_RERUN = True` to redo them, or `REBUILD_INDEX = True` to rebuild the PDF indexes.

---

## 5. Launch the dashboard

If the data is already generated, you only need:

```bash
streamlit run app.py
```

Open http://localhost:8501. The sidebar has three pages:

- **Company Intelligence** (`app.py`)
- **Sector Comparison** (`pages/1_Comparison.py`)
- **AI Company Assistant** (`pages/2_Chatbot.py`)

The dashboard reads these files, so run the pipeline (or restore them) first:

- `output/companies_final_financial_sector_ranked.xlsx`
- `tech_matrix_output/NSE_BSE Companies.xlsx`

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `GROQ_API_KEY is not set` | Add it to `.env` and put it in `groq_key.txt`. |
| `ModuleNotFoundError` | Activate `.venv` and re-run `pip install -r requirements.txt`. |
| Playwright / browser errors in `save_urls.py` | Run `playwright install chromium`. |
| `Project folder not found: D:\GENXAI...` | Apply the notebook path fix in section 3. |
| `Worksheet named 'All Companies' not found` | Rename the sheet, or enable `RUN_COMPANY_GENERATION`. |
| Many PDF downloads fail | Some company sites block scripted downloads. Re-run `download_urls.py` (it skips finished files) or add the PDFs manually to `annual_reports/` as `<Company Name>.pdf`. |
| NSE/BSE fetch fails or returns empty | The sites rate-limit and change formats. Retry later. For BSE you can place a `BSE_Equity.csv` in the project root as a fallback. |
| CUDA out-of-memory in the notebook | Lower the batch size in the embedder, or set `USE_RERANKER = False`. |
| Dashboard shows "Unable to load project data" | Check that the three data files listed in section 5 exist and that the paths in `utils/data_loader.py` point to them. |
| Port 8501 busy | `streamlit run app.py --server.port 8502` |

---

## Notes

- Scraping NSE/BSE and company websites is subject to their terms of use. Keep the built-in delays.
- LLM extraction costs depend on your Groq plan and rate limits.
- Annual report PDFs and generated Excel files can be large; keep them out of version control.