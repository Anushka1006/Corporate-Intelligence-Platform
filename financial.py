from pathlib import Path
from datetime import datetime
from io import BytesIO
from urllib.parse import urljoin
import re
import shutil
import time
import warnings
import xml.etree.ElementTree as ET

import pandas as pd
import requests
from rapidfuzz import fuzz

try:
    import yfinance as yf
except ImportError:
    yf = None

warnings.filterwarnings("ignore")
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

INPUT_FILE = BASE_DIR / "NSE_BSE_Top_600s_Companies.xlsx"
OUTPUT_FILE = OUTPUT_DIR / "companies_final_financial_sector_ranked.xlsx"
OUTPUT_SHEET = "Financial Data"
MARKET_DATA_ONLY = False
REFRESH_ALL_MARKET_DATA = False


START_COMPANY = 1   
END_COMPANY = None
ONLY_COMPANIES = []
SYMBOL_OVERRIDES = {}

MAX_RETRIES = 3
REQUEST_DELAY = 1.5
REQUEST_TIMEOUT = 45

MIN_MATCH_SCORE = 90
MIN_MATCH_MARGIN = 8

USE_YFINANCE_FALLBACK = True

USE_BSE_FALLBACK = True
BSE_LIST_URL = ("https://api.bseindia.com/BseIndiaAPI/api/ListofScripData/w"
                "?Group=&Scripcode=&industry=&segment=Equity&status=Active")
BSE_LIST_FILE = BASE_DIR / "BSE_Equity.csv"

NSE_MASTER_URL = "https://archives.nseindia.com/content/equities/EQUITY_L.csv"
NSE_HOME = "https://www.nseindia.com"
NSE_FINANCIAL_RESULTS_URL = NSE_HOME + "/api/corporates-financial-results"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json,text/plain,*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": NSE_HOME + "/",
    "Connection": "keep-alive",
}

OUTPUT_COLUMNS = [
    "Sector Group", "Company_Name", "Ticker", "Exchange",
    "Exchange Company Name", "LTP",
    "Revenue (Billion)", "Revenue Growth", "Net Revenue (Billion)",
    "EBITDA ", "EBITDA Margin", "EPS", "PAT ",
    "EPS (FY-26)", "PAT (FY-26)", "ROE",
    "Debt (Billion)", "Cash (Billion)", "Debt Equity Ratio",
    "Financial Year", "Updated At",
    "Sector", "Industry", "Market Cap Currency", "Market Cap",
    "Sector Rank", "Sector Leader", "Sector Company Count",
    "Global Market Cap Rank",
]

# ============================================================
# 2. SESSION
# ============================================================

session = requests.Session()
session.headers.update(HEADERS)


def initialize_nse_session():
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            session.get(NSE_HOME, timeout=REQUEST_TIMEOUT).raise_for_status()
            session.get(NSE_HOME + "/get-quotes/equity?symbol=TCS",
                        timeout=REQUEST_TIMEOUT)
            print("[NSE] Session initialized.")
            return True
        except Exception as error:
            print(f"[NSE SESSION ERROR] Attempt {attempt}: {error}")
            time.sleep(2 * attempt)
    return False


# ============================================================
# 3. HELPERS
# ============================================================

def clean_text(value):
    if value is None or pd.isna(value):
        return ""
    value = str(value).upper().strip().replace("&", " AND ")
    value = re.sub(r"[^A-Z0-9 ]", " ", value)
    return re.sub(r"\s+", " ", value).strip()


ABBREVIATIONS = {
    "CO": "COMPANY", "CORP": "CORPORATION", "INDS": "INDUSTRIES",
    "INFRA": "INFRASTRUCTURE", "INFR": "INFRASTRUCTURE",
    "INV": "INVESTMENT", "INVT": "INVESTMENT",
    "PHARMACEUTICALS": "PHARMA", "PHARMACEUTICAL": "PHARMA",
    "TECH": "TECHNOLOGIES", "TECHNOLOGY": "TECHNOLOGIES",
}


def normalize_company_name(name):
    suffixes = {"LIMITED", "LTD", "PRIVATE", "PVT", "PUBLIC",
                "PLC", "INCORPORATED", "INC"}
    raw = "" if name is None or pd.isna(name) else str(name).upper()
    raw = re.sub(r"\(\s*I\s*\)", " INDIA ", raw)       # "(I)" -> INDIA
    words = clean_text(raw).split()
    if words and words[0] == "THE":                       # "The Federal Bank"
        words = words[1:]
    words = [ABBREVIATIONS.get(w, w) for w in words]
    while words and words[-1] in suffixes:
        words.pop()
    return " ".join(words)


def valid_company_name(value):
    value = str(value).strip()
    return (len(value) >= 3 and bool(re.search(r"[A-Za-z]", value))
            and value.lower() != "nan")


def safe_float(value):
    try:
        if value is None:
            return None
        if isinstance(value, str):
            value = value.replace(",", "").strip()
        result = float(value)
        return None if pd.isna(result) else result
    except (TypeError, ValueError):
        return None


def first_valid(*values):
    for value in values:
        if value is None:
            continue
        try:
            if pd.isna(value):
                continue
        except (TypeError, ValueError):
            pass
        if isinstance(value, str) and not value.strip():
            continue
        return value
    return None


def to_billion(value):
    value = safe_float(value)
    return None if value is None else value / 1e9


def ratio(numerator, denominator):
    n, d = safe_float(numerator), safe_float(denominator)
    if n is None or d is None or d == 0:
        return None
    return n / d


def calculate_growth(current, previous):
    c, p = safe_float(current), safe_float(previous)
    if c is None or p is None or p == 0:
        return None
    return (c - p) / abs(p)


def name_similarity(a, b):
    a, b = normalize_company_name(a), normalize_company_name(b)
    if not a or not b:
        return 0
    return max(fuzz.ratio(a, b), fuzz.token_sort_ratio(a, b))


def get_input_value(row, columns):
    for column in columns:
        if column in row.index:
            value = row.get(column)
            if value is not None and not pd.isna(value) and str(value).strip():
                return value
    return None


# ============================================================
# 4. NSE MASTER + MATCHING
# ============================================================

def download_nse_master():
    print("\n[NSE MASTER] Downloading official NSE securities list...")
    response = session.get(NSE_MASTER_URL, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()

    df = pd.read_csv(BytesIO(response.content), dtype=str)
    df.columns = [str(c).strip().upper() for c in df.columns]

    required = {"SYMBOL", "NAME OF COMPANY", "SERIES", "ISIN NUMBER"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"NSE master missing columns: {missing}. "
                         f"Received: {list(df.columns)}")

    df = df[df["SERIES"].astype(str).str.strip().str.upper().eq("EQ")].copy()
    df["Normalized_Name"] = df["NAME OF COMPANY"].apply(normalize_company_name)
    df["SYMBOL"] = df["SYMBOL"].astype(str).str.strip().str.upper()
    df["ISIN NUMBER"] = df["ISIN NUMBER"].astype(str).str.strip().str.upper()
    df = df.drop_duplicates(subset=["SYMBOL"]).copy()
    df = df[df["Normalized_Name"].ne("")].copy()

    df["EXCHANGE"] = "NSE"
    df["BSE_ID"] = ""

    print(f"[NSE MASTER] Loaded {len(df)} equity companies.")
    return df


def download_bse_master():
    """BSE scrip list in the same schema as the NSE master
    (SYMBOL = BSE scrip code, NAME OF COMPANY = issuer name)."""
    if not USE_BSE_FALLBACK:
        return None
    print("\n[BSE MASTER] Loading BSE equity list...")
    try:
        if BSE_LIST_FILE.exists():
            df = pd.read_csv(BSE_LIST_FILE, dtype=str)
            print(f"[BSE MASTER] Using local file {BSE_LIST_FILE.name}")
        else:
            r = requests.get(
                BSE_LIST_URL,
                headers={"User-Agent": HEADERS["User-Agent"],
                         "Accept": "application/json, text/plain, */*",
                         "Referer": "https://www.bseindia.com/",
                         "Origin": "https://www.bseindia.com"},
                timeout=REQUEST_TIMEOUT)
            r.raise_for_status()
            df = pd.DataFrame(r.json())
    except Exception as error:
        print(f"[BSE MASTER] Could not load: {error}\n"
              f"             BSE fallback disabled. See BSE_LIST_FILE note in config.")
        return None

    df.columns = [str(c).strip().upper() for c in df.columns]

    def col(*names):
        return next((n for n in names if n in df.columns), None)

    code = col("SCRIP_CD", "SECURITY CODE")
    name = col("ISSUER_NAME", "ISSUER NAME", "SCRIP_NAME", "SECURITY NAME")
    sid = col("SCRIP_ID", "SECURITY ID")
    isin = col("ISIN_NUMBER", "ISIN NO", "ISIN NO.")
    if not code or not name:
        print(f"[BSE MASTER] Unexpected columns {list(df.columns)}; BSE fallback disabled.")
        return None

    out = pd.DataFrame({
        "SYMBOL": df[code].astype(str).str.strip(),
        "NAME OF COMPANY": df[name].astype(str).str.strip(),
        "BSE_ID": df[sid].astype(str).str.strip() if sid else "",
        "ISIN NUMBER": df[isin].astype(str).str.strip().str.upper() if isin else "",
    })
    out["EXCHANGE"] = "BSE"
    out["Normalized_Name"] = out["NAME OF COMPANY"].apply(normalize_company_name)
    out = out.drop_duplicates(subset=["SYMBOL"])
    out = out[out["Normalized_Name"].ne("")].copy()
    print(f"[BSE MASTER] Loaded {len(out)} BSE equity companies.")
    return out


def _pack(best):
    return {"Symbol": best["SYMBOL"], "Master_Name": best["NAME OF COMPANY"],
            "Exchange": best.get("EXCHANGE", "NSE"),
            "Aux": best.get("BSE_ID", "")}


OVERRIDES_NORM = {clean_text(k): v for k, v in SYMBOL_OVERRIDES.items()}


def resolve_company(company_name, master, input_row=None, use_overrides=True):
    """Returns (match_dict_or_None, reason)."""
    # 0) manual override
    sym = OVERRIDES_NORM.get(clean_text(company_name)) if use_overrides else None
    if sym:
        hit = master[master["SYMBOL"] == sym.upper()]
        if hit.empty:
            return None, f"Override symbol {sym} not found in NSE EQ master"
        print(f"[OVERRIDE] {company_name} -> {sym} "
              f"(NSE name: {hit.iloc[0]['NAME OF COMPANY']})  <-- verify")
        return _pack(hit.iloc[0]), "override"

    # 1) ISIN / ticker supplied in the input file
    if input_row is not None:
        isin = get_input_value(input_row, ["ISIN", "Isin", "isin"])
        if isin:
            hit = master[master["ISIN NUMBER"] == str(isin).strip().upper()]
            if not hit.empty:
                return _pack(hit.iloc[0]), "isin"
        tkr = get_input_value(input_row, ["Ticker", "Symbol", "NSE Symbol"])
        if tkr:
            hit = master[master["SYMBOL"] == str(tkr).strip().upper()]
            if not hit.empty:
                return _pack(hit.iloc[0]), "ticker"

    normalized = normalize_company_name(company_name)
    if not normalized:
        return None, "Empty name after normalisation"

    # 2) exact normalised name
    exact = master[master["Normalized_Name"] == normalized]
    if len(exact) == 1:
        return _pack(exact.iloc[0]), "exact"
    if len(exact) > 1:
        names = "; ".join(exact["SYMBOL"].head(4))
        return None, f"Several exact matches ({names}) - add an override"

    # 3) input is a whole-word prefix of exactly one NSE name
    prefix = master[master["Normalized_Name"].str.startswith(normalized + " ")]
    if len(prefix) == 1:
        return _pack(prefix.iloc[0]), "prefix"

    # 4) fuzzy with margin check
    scored = master.copy()
    scored["Match_Score"] = scored["Normalized_Name"].apply(
        lambda c: name_similarity(normalized, c))
    scored = scored.sort_values("Match_Score", ascending=False)
    top = scored.head(3)
    suggestions = "; ".join(
        f"{r['SYMBOL']} ({r['NAME OF COMPANY']}, {r['Match_Score']:.0f})"
        for _, r in top.iterrows())

    good = scored[scored["Match_Score"] >= MIN_MATCH_SCORE]
    if good.empty:
        print(f"[NO NSE MATCH] {company_name}. Closest: {suggestions}")
        return None, f"No match >= {MIN_MATCH_SCORE}. Closest: {suggestions}"

    best = good.iloc[0]
    if len(good) > 1:
        second = good.iloc[1]
        if best["Match_Score"] - second["Match_Score"] < MIN_MATCH_MARGIN:
            print(f"[AMBIGUOUS MATCH] {company_name}: {suggestions}")
            return None, f"Ambiguous: {suggestions}"

    return _pack(best), f"fuzzy {best['Match_Score']:.0f}"


# ============================================================
# 5. FILINGS METADATA
# ============================================================

def nse_get_json(url, params):
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.get(url, params=params, timeout=REQUEST_TIMEOUT)
            response.raise_for_status()
            return response.json()
        except Exception as error:
            print(f"[NSE ERROR] {params.get('symbol')}: {error}")
            if attempt < MAX_RETRIES:
                initialize_nse_session()
                time.sleep(REQUEST_DELAY * attempt)
    return None


def get_annual_filings(symbol):
    """Returns (filings, basis) preferring consolidated, else standalone."""
    data = nse_get_json(
        NSE_FINANCIAL_RESULTS_URL,
        {"index": "equities", "symbol": symbol, "period": "Annual"},
    )
    if data is None:
        return [], None

    if isinstance(data, dict):
        data = data.get("data") or data.get("results") or []
    if not isinstance(data, list):
        return [], None

    filings = [f for f in data if isinstance(f, dict)]

    def is_consolidated(f):
        v = str(f.get("consolidated", "")).strip().lower()
        return v in {"consolidated", "yes", "true", "y", "c"}

    def period_end(f):
        value = f.get("toDate") or f.get("periodEnd") or f.get("period") or ""
        parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
        return parsed if not pd.isna(parsed) else pd.Timestamp.min

    consolidated = sorted([f for f in filings if is_consolidated(f)],
                          key=period_end, reverse=True)
    standalone = sorted([f for f in filings if not is_consolidated(f)],
                        key=period_end, reverse=True)

    print(f"[FILINGS] {symbol}: {len(consolidated)} consolidated, "
          f"{len(standalone)} standalone annual filings")

    if consolidated:
        return consolidated, "Consolidated"
    if standalone:
        return standalone, "Standalone"
    return [], None


def get_xbrl_url(filing):
    for field in ["xbrl", "xbrlFile", "xbrlUrl", "xbrlURL", "xbrl_file"]:
        value = filing.get(field)
        if isinstance(value, str) and value.strip().lower().startswith(
                ("http", "/")):
            value = value.strip()
            return value if value.startswith("http") else urljoin(NSE_HOME, value)
    return None


def download_xbrl(url):
    try:
        response = session.get(url, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        return ET.fromstring(response.content)
    except Exception as error:
        print(f"[XBRL DOWNLOAD ERROR] {error}")
        return None


# ============================================================
# 6. XBRL PARSING
# ============================================================

def local_name(tag):
    if "}" in tag:
        tag = tag.split("}", 1)[1]
    return tag.split(":")[-1]


def parse_contexts(root):
    """Skip dimensional (segment/scenario) contexts so segment figures are
    never mistaken for company totals."""
    contexts = {}
    for el in root.iter():
        if local_name(el.tag) != "context":
            continue
        cid = el.attrib.get("id")
        if not cid:
            continue

        start = end = instant = None
        has_dimension = False
        for child in el.iter():
            name = local_name(child.tag)
            if name in {"segment", "scenario"}:
                has_dimension = True
            elif name == "startDate":
                start = (child.text or "").strip()
            elif name == "endDate":
                end = (child.text or "").strip()
            elif name == "instant":
                instant = (child.text or "").strip()

        if not has_dimension:
            contexts[cid] = {"start": start, "end": end, "instant": instant}
    return contexts


def annual_periods(contexts):
    """[(end_date, duration_ctx_id, instant_ctx_id), ...] newest first."""
    periods = {}
    for cid, c in contexts.items():
        if c["start"] and c["end"]:
            try:
                days = (pd.Timestamp(c["end"]) - pd.Timestamp(c["start"])).days
            except Exception:
                continue
            if 300 <= days <= 400:
                periods.setdefault(c["end"], {})["dur"] = cid

    for cid, c in contexts.items():
        if c["instant"] and c["instant"] in periods:
            periods[c["instant"]]["inst"] = cid

    return sorted(
        [(end, v.get("dur"), v.get("inst")) for end, v in periods.items()],
        key=lambda x: x[0], reverse=True)


def extract_facts(root):
    """facts[name][context] = float ; also returns the declared scale."""
    facts, scale_text = {}, ""
    for el in root.iter():
        name = local_name(el.tag).lower()
        text = (el.text or "").strip()
        if name == "scaleoffigures" and text:
            scale_text = text.lower()
        ctx = el.attrib.get("contextRef")
        if ctx is None or not text:
            continue
        value = safe_float(text)
        if value is None:
            continue
        facts.setdefault(name, {}).setdefault(ctx, value)

    if "crore" in scale_text:
        scale = 1e7
    elif "lakh" in scale_text or "lac" in scale_text:
        scale = 1e5
    elif "million" in scale_text:
        scale = 1e6
    elif "thousand" in scale_text:
        scale = 1e3
    else:
        scale = 1.0
    return facts, scale


ALIASES = {
    "revenue": ["revenuefromoperations", "revenuefromoperationsfromcontinuingoperations",
                "revenue", "totalrevenue", "revenuefromoperationsnetofexciseduty",
                "revenuefromsaleofproducts", "turnovergross"],
    "total_income": ["income", "totalincome", "totalrevenueincludingotherincome"],
    "pbt": ["profitbeforetax", "profitlossbeforetax", "profitbeforeexceptionalitemsandtax"],
    "pat": ["profitlossforperiod", "profitloss", "profitlossfortheperiod",
            "profitlossfromcontinuingoperations", "profitaftertax",
            "profitlossaftertax"],
    "finance_cost": ["financecosts", "financecost", "interestexpense"],
    "depreciation": ["depreciationdepletionandamortisationexpense",
                     "depreciationandamortisationexpense", "depreciationandamortisation"],
    "ebitda_tag": ["ebitda", "earningsbeforeinteresttaxdepreciationandamortisation"],
    "eps_diluted": ["dilutedearningslosspershare",
                    "dilutedearningslosspersharefromcontinuingandanddiscontinuedoperations",
                    "dilutedearningslosspersharefromcontinuingdiscontinuedoperations",
                    "earningspersharediluted", "dilutedearningspershare"],
    "eps_basic": ["basicearningslosspershare",
                  "basicearningslosspersharefromcontinuingandanddiscontinuedoperations",
                  "basicearningslosspersharefromcontinuingdiscontinuedoperations",
                  "earningspersharebasic", "basicearningspershare"],
    # balance sheet (instant)
    "debt": ["borrowings", "totalborrowings", "totaldebt"],
    "debt_parts": ["borrowingsnoncurrent", "noncurrentborrowings", "longtermborrowings",
                   "borrowingscurrent", "currentborrowings", "shorttermborrowings"],
    "cash": ["cashandcashequivalents", "cashandcashequivalentsandbankbalances",
             "cashcashequivalentsandbankbalances", "cashandbankbalances"],
    "equity": ["equity", "totalequity", "equityattributabletoownersofparent"],
}

PER_SHARE = {"eps_basic", "eps_diluted"}
BALANCE = {"debt", "cash", "equity"}


def pick(facts, key, ctx):
    if ctx is None:
        return None
    for alias in ALIASES[key]:
        value = facts.get(alias, {}).get(ctx)
        if value is not None:
            return value
    return None


def sum_debt_parts(facts, ctx):
    if ctx is None:
        return None
    total, found = 0.0, False
    for alias in ALIASES["debt_parts"]:
        v = facts.get(alias, {}).get(ctx)
        if v is not None:
            total += v
            found = True
    return total if found else None


def build_period_values(facts, scale, dur_ctx, inst_ctx):
    def scaled(key, ctx):
        v = pick(facts, key, ctx)
        return None if v is None else v * (1.0 if key in PER_SHARE else scale)

    v = {k: scaled(k, dur_ctx) for k in
         ["revenue", "total_income", "pbt", "pat", "finance_cost",
          "depreciation", "ebitda_tag", "eps_diluted", "eps_basic"]}

    for k in BALANCE:
        v[k] = scaled(k, inst_ctx)

    if v["debt"] is None:
        parts = sum_debt_parts(facts, inst_ctx)
        v["debt"] = None if parts is None else parts * scale

    if v["revenue"] is None:
        v["revenue"] = v["total_income"]

    ebitda = v["ebitda_tag"]
    if ebitda is None and None not in (v["pbt"], v["finance_cost"],
                                       v["depreciation"]):
        ebitda = v["pbt"] + v["finance_cost"] + v["depreciation"]
    v["ebitda"] = ebitda

    v["roe"] = ratio(v["pat"], v["equity"])          # on closing equity
    v["debt_equity"] = ratio(v["debt"], v["equity"])
    v["eps"] = first_valid(v["eps_diluted"], v["eps_basic"])
    return v


def extract_from_xbrl(root):
    """Returns (latest, previous); previous may be None."""
    periods = annual_periods(parse_contexts(root))
    if not periods:
        return None, None

    facts, scale = extract_facts(root)

    end, dur, inst = periods[0]
    latest = build_period_values(facts, scale, dur, inst)
    latest["financial_year"] = end

    previous = None
    if len(periods) > 1:
        p_end, p_dur, p_inst = periods[1]
        gap = (pd.Timestamp(end) - pd.Timestamp(p_end)).days
        if 330 <= gap <= 400:
            previous = build_period_values(facts, scale, p_dur, p_inst)
            previous["financial_year"] = p_end

    return latest, previous


# ============================================================
# 7. YFINANCE FALLBACK (financials only)
# ============================================================

def _yf_row(df, labels, col):
    for label in labels:
        if label in df.index:
            v = safe_float(df.at[label, col])
            if v is not None:
                return v
    return None


def yfinance_values(symbol, suffix=".NS"):
    """symbol may be a string or a list of candidate tickers."""
    if yf is None or not USE_YFINANCE_FALLBACK:
        return None, None

    candidates = symbol if isinstance(symbol, list) else [symbol]
    for sym in [c for c in candidates if c]:
        try:
            t = yf.Ticker(str(sym) + suffix)
            fin, bs = t.financials, t.balance_sheet
            if fin is None or fin.empty:
                continue

            def build(col):
                pat = _yf_row(fin, ["Net Income", "Net Income Common Stockholders"], col)
                v = {
                    "revenue": _yf_row(fin, ["Total Revenue", "Operating Revenue"], col),
                    "pat": pat,
                    "ebitda": _yf_row(fin, ["EBITDA", "Normalized EBITDA"], col),
                    "eps": _yf_row(fin, ["Diluted EPS", "Basic EPS"], col),
                    "debt": None, "cash": None, "roe": None, "debt_equity": None,
                    "financial_year": pd.Timestamp(col).strftime("%Y-%m-%d"),
                }
                if bs is not None and not bs.empty and col in bs.columns:
                    equity = _yf_row(bs, ["Stockholders Equity", "Common Stock Equity"], col)
                    v["debt"] = _yf_row(bs, ["Total Debt"], col)
                    v["cash"] = _yf_row(bs, ["Cash And Cash Equivalents"], col)
                    v["roe"] = ratio(pat, equity)
                    v["debt_equity"] = ratio(v["debt"], equity)
                return v

            cols = list(fin.columns)
            return build(cols[0]), (build(cols[1]) if len(cols) > 1 else None)
        except Exception as error:
            print(f"[YFINANCE ERROR] {sym}{suffix}: {error}")
    return None, None


FILL_KEYS = ["revenue", "pat", "ebitda", "eps", "debt", "cash",
             "roe", "debt_equity"]


# ============================================================
# 8. ANNUAL RESULTS ORCHESTRATION
# ============================================================

def extract_annual_results(symbol):
    filings, basis = get_annual_filings(symbol)

    for idx, filing in enumerate(filings[:4]):
        url = get_xbrl_url(filing)
        if not url:
            continue
        root = download_xbrl(url)
        time.sleep(REQUEST_DELAY)
        if root is None:
            continue

        latest, previous = extract_from_xbrl(root)
        if latest is None:
            continue

        # previous year not inside the same file -> try the next filing
        if previous is None and idx + 1 < len(filings):
            nxt_url = get_xbrl_url(filings[idx + 1])
            nxt_root = download_xbrl(nxt_url) if nxt_url else None
            if nxt_root is not None:
                nxt_latest, _ = extract_from_xbrl(nxt_root)
                if nxt_latest:
                    gap = (pd.Timestamp(latest["financial_year"])
                           - pd.Timestamp(nxt_latest["financial_year"])).days
                    if 330 <= gap <= 400:
                        previous = nxt_latest

        return {"latest": latest, "previous": previous, "basis": basis}

    return {"latest": None, "previous": None, "basis": None}


def merge_with_fallback(symbol, annual):
    """Fill gaps from yfinance (same fiscal year-end only), or use it
    entirely when XBRL gave nothing."""
    yf_latest, yf_prev = yfinance_values(symbol)

    if annual["latest"] is None:
        if yf_latest is not None:
            annual.update(latest=yf_latest, previous=yf_prev)
        return annual

    if yf_latest is None:
        return annual

    if yf_latest["financial_year"] == annual["latest"]["financial_year"]:
        for k in FILL_KEYS:
            if annual["latest"].get(k) is None and yf_latest.get(k) is not None:
                annual["latest"][k] = yf_latest[k]
        if annual["previous"] is None and yf_prev is not None:
            annual["previous"] = yf_prev
    return annual


# ============================================================
# 8b. MARKET DATA (Yahoo): LTP, Market Cap, Sector, Industry
# ============================================================

def _fast(fast, *keys):
    for key in keys:
        try:
            value = safe_float(fast[key])
            if value is not None:
                return value
        except Exception:
            continue
    return None


def fetch_market_data(exchange, symbol, aux=None):
    """LTP, market cap, currency, sector, industry from Yahoo Finance."""
    if yf is None:
        print("[MARKET DATA] yfinance not installed (pip install yfinance)")
        return None

    suffix = ".NS" if str(exchange).upper() == "NSE" else ".BO"
    tickers = [f"{symbol}{suffix}"]
    if suffix == ".BO" and aux:
        tickers.append(f"{aux}{suffix}")

    for ticker in tickers:
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                stock = yf.Ticker(ticker)

                try:
                    fast = stock.fast_info
                except Exception:
                    fast = {}
                try:
                    info = stock.get_info()
                    if not isinstance(info, dict):
                        info = {}
                except Exception:
                    info = {}

                ltp = first_valid(
                    _fast(fast, "last_price", "lastPrice"),
                    safe_float(info.get("currentPrice")),
                    safe_float(info.get("regularMarketPrice")),
                    safe_float(info.get("previousClose")))

                market_cap = first_valid(
                    _fast(fast, "market_cap", "marketCap"),
                    safe_float(info.get("marketCap")))
                if market_cap is None and ltp is not None:
                    shares = first_valid(_fast(fast, "shares"),
                                         safe_float(info.get("sharesOutstanding")))
                    if shares is not None:
                        market_cap = ltp * shares

                sector = info.get("sector") or info.get("sectorDisp")
                industry = info.get("industry") or info.get("industryDisp")

                if ltp is None and market_cap is None and not sector:
                    raise ValueError("Yahoo returned no market data")

                return {
                    "LTP": ltp,
                    "Market Cap": market_cap,
                    "Currency": first_valid(info.get("currency"), "INR"),
                    "Sector": sector or "Unclassified",
                    "Industry": industry or "Unclassified",
                }
            except Exception as error:
                print(f"[MARKET DATA] {ticker} attempt {attempt}: {error}")
                time.sleep(REQUEST_DELAY * attempt)
    return None


def rank_and_sort(df):
    """Rank companies within sectors and order sectors by their leader's cap."""
    df = df.copy()
    if df.empty:
        return df

    df["Sector"] = (
        df["Sector"].fillna("").astype(str).str.strip()
        .replace({"": "Unclassified", "nan": "Unclassified", "None": "Unclassified"})
    )
    df["Market Cap"] = pd.to_numeric(
        df["Market Cap"].astype(str).str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False).str.strip(),
        errors="coerce",
    )
    df.loc[df["Market Cap"] <= 0, "Market Cap"] = pd.NA

    # Sequential ranks: 1, 2, 3...; tied market caps are ordered by company name.
    df = df.sort_values(
        ["Sector", "Market Cap", "Company_Name"],
        ascending=[True, False, True], na_position="last", kind="stable"
    ).reset_index(drop=True)

    df["Sector Rank"] = (
        df.groupby("Sector", sort=False)["Market Cap"]
        .rank(method="first", ascending=False, na_option="keep")
        .astype("Int64")
    )
    df["Sector Company Count"] = (
        df.groupby("Sector")["Company_Name"].transform("count").astype("Int64")
    )

    leaders = (
        df[df["Market Cap"].notna()]
        .sort_values(["Sector", "Market Cap", "Company_Name"],
                     ascending=[True, False, True], kind="stable")
        .drop_duplicates(subset=["Sector"])
        .set_index("Sector")["Company_Name"].to_dict()
    )
    df["Sector Leader"] = df["Sector"].map(leaders)

    # Global rank uses competition ranking; rows without market cap stay blank.
    df["Global Market Cap Rank"] = (
        df["Market Cap"].rank(method="min", ascending=False, na_option="keep")
        .astype("Int64")
    )

    # Put the strongest sector first, based on its highest company market cap.
    sector_order = (
        df.groupby("Sector")["Market Cap"].max()
        .sort_values(ascending=False, na_position="last", kind="stable")
    )
    order_map = {sector: i for i, sector in enumerate(sector_order.index)}
    df["_sector_order"] = df["Sector"].map(order_map)
    df = df.sort_values(
        ["_sector_order", "Market Cap", "Company_Name"],
        ascending=[True, False, True], na_position="last", kind="stable"
    ).drop(columns="_sector_order").reset_index(drop=True)
    return df


# ============================================================
# 9. BUILD COMPANY ROW
# ============================================================

def fetch_company(company_name, input_row, master, used_symbols, bse_master=None):
    print(f"\n[COMPANY] {company_name}")

    resolved, reason = resolve_company(company_name, master, input_row)
    if resolved is None and bse_master is not None:
        nse_reason = reason
        resolved, bse_reason = resolve_company(company_name, bse_master,
                                               input_row, use_overrides=False)
        if resolved is None:
            return None, f"NSE: {nse_reason} | BSE: {bse_reason}"
        print(f"[BSE] Not on NSE; matched on BSE -> code {resolved['Symbol']} "
              f"({resolved['Master_Name']})  <-- verify")
    if resolved is None:
        return None, reason

    exchange = resolved["Exchange"]
    symbol = resolved["Symbol"]
    key = f"{exchange}:{symbol}"
    if key in used_symbols:
        print(f"[DUPLICATE SYMBOL] {key} already assigned.")
        return None, f"Duplicate: {key} already used by another row"

    print(f"[MATCHED] {company_name} -> {symbol}")

    if exchange == "BSE":
        # no free official BSE XBRL feed -> Yahoo data for <code>.BO
        candidates = [symbol] + ([resolved["Aux"]] if resolved.get("Aux") else [])
        y_latest, y_prev = yfinance_values(candidates, ".BO")
        annual = {"latest": y_latest, "previous": y_prev, "basis": "yfinance BSE"}
    else:
        annual = merge_with_fallback(symbol, extract_annual_results(symbol))
    latest = annual["latest"] or {}
    previous = annual["previous"] or {}

    if not latest:
        print(f"[FINANCIALS UNAVAILABLE] {symbol} (row written with blanks)")

    revenue = latest.get("revenue")
    ebitda = latest.get("ebitda")
    fy_raw = latest.get("financial_year")

    try:
        fy_ts = pd.Timestamp(fy_raw)
        financial_year = fy_ts.strftime("%d-%m-%Y")
        is_fy26 = fy_ts.year == 2026
    except Exception:
        financial_year, is_fy26 = fy_raw, False

    eps = latest.get("eps")
    pat = latest.get("pat")

    used_symbols.add(key)

    mkt = fetch_market_data(exchange, symbol, resolved.get("Aux")) or {}

    result = {
        "Sector Group": get_input_value(input_row, ["Sector Group", "Sector_Group"]),
        "Company_Name": company_name,
        "Ticker": symbol,
        "Exchange": exchange,
        "Exchange Company Name": resolved["Master_Name"],
        "Revenue (Billion)": to_billion(revenue),
        "Revenue Growth": calculate_growth(revenue, previous.get("revenue")),
        "Net Revenue (Billion)": to_billion(revenue),
        "EBITDA ": to_billion(ebitda),
        "EBITDA Margin": ratio(ebitda, revenue),
        "EPS": eps,
        "PAT ": to_billion(pat),
        "EPS (FY-26)": eps if is_fy26 else None,
        "PAT (FY-26)": to_billion(pat) if is_fy26 else None,
        "ROE": latest.get("roe"),
        "Debt (Billion)": to_billion(latest.get("debt")),
        "Cash (Billion)": to_billion(latest.get("cash")),
        "Debt Equity Ratio": latest.get("debt_equity"),
        "Financial Year": financial_year,
        "Updated At": datetime.now().strftime("%d-%m-%Y %H:%M"),
        "LTP": mkt.get("LTP"),
        "Sector": mkt.get("Sector", "Unclassified"),
        "Industry": mkt.get("Industry", "Unclassified"),
        "Market Cap Currency": mkt.get("Currency"),
        "Market Cap": mkt.get("Market Cap"),
    }

    print(f"[DONE] {symbol}: revenue={result['Revenue (Billion)']} bn, "
          f"PAT={result['PAT ']} bn")
    return result, "OK"


# ============================================================
# 10. SAVE / LOAD
# ============================================================

def save_output(results):
    """Save a checkpoint/final result as an Excel workbook."""
    df = pd.DataFrame(results)
    for column in OUTPUT_COLUMNS:
        if column not in df.columns:
            df[column] = None
    df = rank_and_sort(df[OUTPUT_COLUMNS])

    # Excel cannot reliably store pandas nullable integer extension arrays
    # across every pandas/openpyxl version; convert rank/count columns safely.
    for col in ["Sector Rank", "Sector Company Count", "Global Market Cap Rank"]:
        if col in df.columns:
            df[col] = df[col].astype(object).where(df[col].notna(), None)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)
        ws = writer.sheets[OUTPUT_SHEET]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        # Readable widths without making the workbook excessively wide.
        for column_cells in ws.columns:
            letter = column_cells[0].column_letter
            max_len = max((len(str(c.value)) if c.value is not None else 0)
                          for c in column_cells[:200])
            ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 34)
    return df


def _read_existing_file(path):
    if path.suffix.lower() in {".xlsx", ".xls"}:
        try:
            return pd.read_excel(path, sheet_name=OUTPUT_SHEET, dtype={"Ticker": str})
        except ValueError:
            return pd.read_excel(path, dtype={"Ticker": str})
    return pd.read_csv(path, dtype={"Ticker": str}, encoding="utf-8-sig")


def load_existing_rows():
    """
    Load only the existing final output workbook.
    Does not merge any external input files.
    """
    if not OUTPUT_FILE.exists():
        print("[LOAD] No previous output found. Starting fresh.")
        return []

    try:
        df = _read_existing_file(OUTPUT_FILE)

        if "Company" in df.columns and "Company_Name" not in df.columns:
            df = df.rename(columns={"Company": "Company_Name"})

        if "Company Name" in df.columns and "Company_Name" not in df.columns:
            df = df.rename(columns={"Company Name": "Company_Name"})

        for column in OUTPUT_COLUMNS:
            if column not in df.columns:
                df[column] = None

        df["Exchange"] = df["Exchange"].fillna("NSE")
        df["Ticker"] = df["Ticker"].astype(str).str.strip()

        df = df[
            df["Ticker"].ne("") &
            df["Ticker"].ne("nan")
        ]

        df = df.drop_duplicates(
            subset=["Exchange", "Ticker"],
            keep="last"
        )

        print(f"[LOAD] {len(df)} existing companies loaded from final output.")

        return df.to_dict("records")

    except Exception as error:
        print(f"[LOAD WARNING] Could not load existing output: {error}")
        return []
def needs_market_data(rec):
    if REFRESH_ALL_MARKET_DATA:
        return True
    sector = str(rec.get("Sector") or "").strip()
    return (safe_float(rec.get("LTP")) is None
            or safe_float(rec.get("Market Cap")) is None
            or sector in {"", "nan", "Unclassified"})


def backfill_market_data(results):
    todo = [r for r in results if needs_market_data(r)]
    print(f"[BACKFILL] {len(todo)} of {len(results)} rows need market data")
    failed = []

    for count, rec in enumerate(todo, start=1):
        print(f"[{count}/{len(todo)}] {rec.get('Company_Name')} "
              f"({rec.get('Exchange')}:{rec.get('Ticker')})")
        mkt = fetch_market_data(rec.get("Exchange"), str(rec.get("Ticker")))
        if mkt is None:
            failed.append(rec.get("Company_Name"))
        else:
            rec["LTP"] = first_valid(mkt["LTP"], rec.get("LTP"))
            rec["Market Cap"] = first_valid(mkt["Market Cap"], rec.get("Market Cap"))
            rec["Market Cap Currency"] = mkt["Currency"]
            if mkt["Sector"] != "Unclassified":
                rec["Sector"] = mkt["Sector"]
            if mkt["Industry"] != "Unclassified":
                rec["Industry"] = mkt["Industry"]
        if count % 10 == 0:
            save_output(results)           # checkpoint
        time.sleep(REQUEST_DELAY)

    return save_output(results), failed


def print_summary(df, failed=None):
    print("\n" + "=" * 75)
    print(f"Excel output saved: {OUTPUT_FILE} (sheet: {OUTPUT_SHEET})")
    print(f"Total companies: {len(df)}")
    print(f"With market cap: {int(df['Market Cap'].notna().sum())} | "
          f"without: {int(df['Market Cap'].isna().sum())}")
    print(f"Unclassified sector: {int(df['Sector'].eq('Unclassified').sum())}")
    if failed:
        print(f"No Yahoo market data for {len(failed)}: {failed}")
    top = df.sort_values("Market Cap", ascending=False, na_position="last").head(10)
    print("\nTop 10 by market cap:")
    print(top[["Global Market Cap Rank", "Company_Name", "Sector",
               "Market Cap", "Sector Rank"]].to_string(index=False))
    print("=" * 75)


# ============================================================
# 11. MAIN
# ============================================================

def main():
    print("=" * 75)
    print("ONE-SHOT NSE/BSE FINANCIALS + YAHOO MARKET DATA + SECTOR RANKING")
    print("=" * 75)

    results = load_existing_rows()

    print(f"[START] Continuing with {len(results)} existing result rows.")
    print(f"[INPUT] {INPUT_FILE.name}")
    print(f"[OUTPUT] {OUTPUT_FILE.name}")

    # ---------- mode 1: backfill LTP / market cap / sector / industry ----
    if MARKET_DATA_ONLY:
        if not results:
            raise RuntimeError(
                "No existing rows found. Check MERGE_INPUT_FILES / output folder.")
        final_df, failed = backfill_market_data(results)
        print_summary(final_df, failed)
        return

    # ---------- mode 2: normal run (filings + market data) ---------------
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"Input file not found: {INPUT_FILE}")

    if INPUT_FILE.suffix.lower() in {".xlsx", ".xls"}:
        companies = pd.read_excel(INPUT_FILE, dtype=str)
    else:
        companies = pd.read_csv(INPUT_FILE, dtype=str, encoding="utf-8-sig")

    # Accept common variants of the company-name header.
    if "Company_Name" not in companies.columns:
        for candidate in ["Company Name", "Company", "Name"]:
            if candidate in companies.columns:
                companies = companies.rename(columns={candidate: "Company_Name"})
                break
    if "Company_Name" not in companies.columns:
        raise ValueError(
            "Input file must contain a company-name column: Company_Name, "
            "Company Name, Company or Name."
        )

    if ONLY_COMPANIES:
        wanted = {clean_text(n) for n in ONLY_COMPANIES}
        mask = companies["Company_Name"].astype(str).apply(clean_text).isin(wanted)
        selected = companies[mask].copy()
        print(f"RETRY MODE: {len(selected)} of {len(ONLY_COMPANIES)} found")
    else:
        if START_COMPANY < 1 or END_COMPANY < START_COMPANY:
            raise ValueError("Invalid company range.")
        if START_COMPANY > len(companies):
            raise ValueError("START_COMPANY exceeds the input row count.")
        end = min(END_COMPANY, len(companies))
        selected = companies.iloc[START_COMPANY - 1:end].copy()
        print(f"Total: {len(companies)} | range {START_COMPANY}-{end} "
              f"| selected {len(selected)}")

    if yf is None:
        raise ImportError("yfinance is required. Install it with: pip install yfinance")

    if not initialize_nse_session():
        raise RuntimeError("Could not initialize NSE session.")

    master = download_nse_master()
    bse_master = download_bse_master()

    used_symbols, skipped = set(), []

    for count, (_, row) in enumerate(selected.iterrows(), start=1):
        company_name = str(row["Company_Name"]).strip()
        print("\n" + "=" * 75)
        print(f"Processing {count}/{len(selected)}: {company_name}")
        print("=" * 75)

        if not valid_company_name(company_name):
            skipped.append((company_name, "Invalid company name"))
            continue

        try:
            result, reason = fetch_company(company_name, row, master,
                                           used_symbols, bse_master)
            if result is None:
                skipped.append((company_name, reason))
            else:
                # replace any older row for the same security
                results = [r for r in results
                           if not (str(r.get("Exchange")) == result["Exchange"]
                                   and str(r.get("Ticker")) == str(result["Ticker"]))]
                results.append(result)
        except Exception as error:
            skipped.append((company_name, f"Error: {error}"))
            print(f"[COMPANY ERROR] {company_name}: {error}")

        save_output(results)          # checkpoint after every company
        time.sleep(REQUEST_DELAY)

    final_df = save_output(results)
    print_summary(final_df)

    if skipped:
        report = OUTPUT_DIR / "skipped_companies.csv"
        pd.DataFrame(skipped, columns=["Company_Name", "Reason"]).to_csv(
            report, index=False, encoding="utf-8-sig")
        print(f"Skipped {len(skipped)} (reasons in {report}):")
        for name, why in skipped:
            print(f"  - {name}: {why}")


if __name__ == "__main__":
    main()