"""
FEC Donor Lookup Script — Batch Mode
=====================================
Reads a donor list from a CSV or Excel file and pulls federal giving history
for each donor from the FEC's free public API, then applies YOUR eligibility
criteria to flag who's worth considering for outreach.

This script is intentionally split in two:
  - config.json          → things you tweak without touching Python at all
  - evaluate_donor()      → the actual eligibility logic, written in Python
                            because real-world rules usually involve branching
                            ("if X and not Y, but Z overrides...") that doesn't
                            fit cleanly into a config file.

See README.md for full setup instructions and PROMPTS.md for example prompts
you can hand to an AI assistant to customize evaluate_donor() for your own
criteria.

SETUP (do this once):
    1. Get a free FEC API key at: https://api.data.gov/signup/
       (takes ~1 minute, they email it to you instantly)

    2. Install dependencies:
       pip3 install requests pandas openpyxl certifi

    3. Copy config.example.json to config.json and fill in your API key,
       file paths, and eligibility criteria.

    4. Edit evaluate_donor() below to match your actual eligibility rules.
       (The example shipped here is a generic placeholder — see PROMPTS.md.)

    5. Run the script:
       python3 fec_donor_lookup.py

EXPECTED SPREADSHEET COLUMNS (case-insensitive, order doesn't matter):
    First | Last | City | State | Company | Occupation | Address (optional)
    Exact header names are configurable in config.json under "columns".

OUTPUT:
    - fec_results_summary.csv        → one row per donor with consider call
    - fec_results/fec_Last_First.txt → full giving history per donor
"""

import sys
import re
import json
import requests

import pandas as pd
import os
import time
from datetime import datetime, date
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import ssl
import certifi
import socket

# Fix for Mac SSL certificate verification error
ssl._create_default_https_context = lambda: ssl.create_default_context(cafile=certifi.where())

# Global socket timeout — catches hangs that requests timeout can miss
socket.setdefaulttimeout(30)


# -------------------------------------------------------
# LOAD CONFIG
# Everything tunable lives in config.json. See config.example.json
# for the full list of options and what each one does.
# -------------------------------------------------------
CONFIG_PATH = "config.json"

def load_config(path):
    if not os.path.exists(path):
        print(f"ERROR: Could not find {path}.")
        print(f"Copy config.example.json to {path} and fill in your own values.")
        sys.exit(1)
    with open(path, "r") as f:
        return json.load(f)

_CONFIG = load_config(CONFIG_PATH)

API_KEY               = _CONFIG["api_key"]
INPUT_FILE             = _CONFIG["input_file"]
OUTPUT_FOLDER          = _CONFIG["output_folder"]
SUMMARY_CSV            = _CONFIG.get("summary_csv", "fec_results_summary.csv")
MIN_DATE               = _CONFIG.get("min_date", "2016-01-01")
MAX_WORKERS            = _CONFIG.get("max_workers", 15)
MAX_RESULTS_PER_DONOR  = _CONFIG.get("max_results_per_donor", 30)

_cols = _CONFIG.get("columns", {})
COL_FIRST      = _cols.get("first", "First")
COL_LAST       = _cols.get("last", "Last")
COL_CITY       = _cols.get("city", "City")
COL_STATE      = _cols.get("state", "State")
COL_COMPANY    = _cols.get("company", "Company")
COL_OCCUPATION = _cols.get("occupation", "Occupation")
COL_ADDRESS    = _cols.get("address", "Address")

BASE_URL = "https://api.open.fec.gov/v1/schedules/schedule_a/"

# -------------------------------------------------------
# BULK DATA CACHE FOLDER
# FEC bulk files are downloaded here once and reused.
# -------------------------------------------------------
BULK_CACHE_FOLDER = "fec_bulk_cache"

# Global in-memory index: {(last_name, first_name, city, state): [records]}
# Built once at startup from cached CSVs, read-only during parallel processing.
_BULK_INDEX = {}

# Thread-safe print lock so terminal output doesn't get garbled
_print_lock = threading.Lock()


# =========================================================
# ELIGIBILITY CRITERIA
# =========================================================
# These come from config.json under "eligibility". Pulled into plain
# variables here so the logic below reads cleanly. If you add new
# criteria fields, add them to config.json AND read them here.
# =========================================================
_elig = _CONFIG.get("eligibility", {})

MIN_SINGLE_DONATION = _elig.get("min_single_donation", 1000)

DISQUALIFYING_EMPLOYERS = [e.lower() for e in _elig.get("disqualifying_employers", [])]

# Keyed by state code (upper-cased), each maps to a list of disqualifying occupations
DISQUALIFYING_OCCUPATIONS_BY_STATE = {
    k.upper(): [v.lower() for v in vals]
    for k, vals in _elig.get("disqualifying_occupations_by_state", {}).items()
}

# Keyed by state code (upper-cased), each maps to a list of occupations needing manual review
VERIFY_OCCUPATIONS_BY_STATE = {
    k.upper(): [v.lower() for v in vals]
    for k, vals in _elig.get("verify_occupations_by_state", {}).items()
}

_special = _elig.get("special_state", {})
SPECIAL_STATE_CODE = (_special.get("code") or "").upper()
SPECIAL_HIGH_TIER_YEARS  = _special.get("min_donations_high_tier_years", 10)
SPECIAL_HIGH_TIER_COUNT  = _special.get("min_donations_high_tier_count", 10)
SPECIAL_MID_TIER_AMOUNT  = _special.get("min_donations_mid_tier_amount", 5000)

_standard = _elig.get("standard_path", {})
STANDARD_HIGH_VALUE_THRESHOLD = _standard.get("high_value_threshold", 20000)
STANDARD_MID_VALUE_THRESHOLD  = _standard.get("mid_value_threshold", 9000)
STANDARD_MID_VALUE_MIN_COUNT  = _standard.get("mid_value_min_count", 3)
STANDARD_RECENCY_YEARS        = _standard.get("recency_years", 5)


def load_donor_list(filepath):
    """Loads the spreadsheet (CSV or Excel) into a list of donor dicts."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(filepath)
    elif ext in (".xlsx", ".xls"):
        df = pd.read_excel(filepath)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Use .csv or .xlsx")

    # Drop unnamed/blank columns that pandas creates from empty CSV columns
    df = df.loc[:, ~df.columns.str.contains(r'^Unnamed', na=False)]
    df.columns = df.columns.str.strip()
    col_map = {c.lower(): c for c in df.columns}

    def get_col(name):
        return col_map.get(name.lower())

    first_col      = get_col(COL_FIRST)
    last_col       = get_col(COL_LAST)
    city_col       = get_col(COL_CITY)
    state_col      = get_col(COL_STATE)
    company_col    = get_col(COL_COMPANY)
    occupation_col = get_col(COL_OCCUPATION)
    address_col    = get_col(COL_ADDRESS)

    if not first_col or not last_col:
        raise ValueError(
            f"Could not find First / Last columns in your file.\n"
            f"Columns found: {list(df.columns)}\n"
            f"Update the \"columns\" section in config.json to match your headers."
        )

    donors = []
    seen = set()
    skipped_blank = 0
    skipped_duplicate = 0

    for _, row in df.iterrows():
        first      = str(row[first_col]).strip()       if first_col       else ""
        last       = str(row[last_col]).strip()        if last_col        else ""
        city       = str(row[city_col]).strip()        if city_col        else ""
        state      = str(row[state_col]).strip()       if state_col       else ""
        company    = str(row[company_col]).strip()     if company_col     else ""
        occupation = str(row[occupation_col]).strip()  if occupation_col  else ""
        address    = str(row[address_col]).strip()     if address_col     else ""

        # Skip rows where First or Last is empty or "nan"
        if not first or not last or first.lower() == "nan" or last.lower() == "nan":
            print(f"  Skipping row — missing First or Last name.")
            skipped_blank += 1
            continue

        # Dedup key: first + last + address + occupation + company (all case-insensitive)
        # Only skips a row if ALL five match — same person entered twice
        dedup_key = (
            first.lower(),
            last.lower(),
            address.lower(),
            occupation.lower(),
            company.lower(),
        )

        if dedup_key in seen:
            print(f"  Skipping duplicate — {first} {last} ({company}, {occupation})")
            skipped_duplicate += 1
            continue

        seen.add(dedup_key)

        donors.append({
            "first_name":  first,
            "last_name":   last,
            "city":        city,
            "state":       state,
            "company":     company,
            "occupation":  occupation,
            "address":     address,
        })

    print(f"Loaded {len(donors)} donors from {filepath}")
    if skipped_blank:
        print(f"  Skipped {skipped_blank} row(s) with missing names.")
    if skipped_duplicate:
        print(f"  Skipped {skipped_duplicate} duplicate row(s).")
    return donors


# Thread-local session so each thread reuses its own HTTP connection
_thread_local = threading.local()

def _get_session():
    """Returns a thread-local requests.Session, creating one if needed."""
    if not hasattr(_thread_local, "session"):
        _thread_local.session = requests.Session()
    return _thread_local.session

# Simple cache to avoid duplicate API calls across threads
# Key: frozenset of params, Value: results list
_api_cache = {}
_api_cache_lock = threading.Lock()

def _get_cached(params):
    key = frozenset((k, str(v)) for k, v in sorted(params.items()) if k != "api_key")
    with _api_cache_lock:
        return _api_cache.get(key)

def _set_cached(params, results):
    key = frozenset((k, str(v)) for k, v in sorted(params.items()) if k != "api_key")
    with _api_cache_lock:
        _api_cache[key] = results


def _fetch_raw(params):
    """
    Core paginated FEC API fetch. Accepts a params dict, returns list of results.
    Handles rate limiting and pagination automatically.
    Times out after 30 seconds per request to prevent hanging.
    Uses persistent HTTP sessions per thread for faster connections.
    Results are cached to avoid duplicate API calls.
    """
    # Check cache first
    cached = _get_cached(params)
    if cached is not None:
        return cached

    session = _get_session()
    all_results = []
    last_indexes = None
    page = 1
    max_retries = 3

    while True:
        p = dict(params)
        if last_indexes:
            p["last_index"] = last_indexes["last_index"]
            p["last_contribution_receipt_amount"] = last_indexes["last_amount"]

        response = None  # reset every page so a failed page can never reuse
                          # a stale response object from a previous page
        for attempt in range(max_retries):
            try:
                response = session.get(BASE_URL, params=p, timeout=30)
                break  # success — exit retry loop
            except requests.exceptions.Timeout:
                response = None
                time.sleep(2)
            except requests.exceptions.RequestException as e:
                response = None
                time.sleep(2)

        if response is None:
            # All retries failed (timeout or connection error) on this page.
            # Fail loudly instead of silently returning thin/duplicated results —
            # raise so the donor gets a visible ERROR row instead of looking
            # like a clean zero-result donor.
            if all_results:
                raise RuntimeError(
                    f"FEC API request failed after {max_retries} retries on page "
                    f"{page} (partial results: {len(all_results)} already fetched). "
                    f"Network/API issue — retry this donor manually."
                )
            else:
                raise RuntimeError(
                    f"FEC API request failed after {max_retries} retries on page "
                    f"{page} (no results fetched yet). Network/API issue — "
                    f"retry this donor manually."
                )

        if response.status_code == 429:
            print("    Rate limited by FEC API — waiting 60 seconds...", flush=True)
            time.sleep(60)
            continue

        if response.status_code != 200:
            break

        data = response.json()
        results = data.get("results", [])
        all_results.extend(results)

        if len(all_results) >= MAX_RESULTS_PER_DONOR:
            break

        if not results or len(results) < 100:
            break

        pagination = data.get("pagination", {})
        last_idx = pagination.get("last_indexes", {})
        if not last_idx.get("last_index"):
            break

        last_indexes = {
            "last_index":  last_idx.get("last_index"),
            "last_amount": last_idx.get("last_contribution_receipt_amount"),
        }
        page += 1

    return all_results


NO_EMPLOYER_BUCKET = {
    "retired", "not employed", "not-employed", "not employed/retired",
    "none", "n/a", "na", "none/retired", "not applicable", "not working",
    "homemaker", "home maker", "housewife", "househusband", "house spouse",
    "unemployed", "student", "disabled", "volunteer",
    "not in workforce", "not in the workforce", "out of workforce",
    "between jobs", "job seeker"
}

# Homemaker and not-employed are used interchangeably in FEC records
HOMEMAKER_EQUIVALENTS = {
    "homemaker", "home maker", "housewife", "househusband",
    "house spouse", "not employed", "not-employed", "not employed/retired",
    "unemployed", "not in workforce", "not in the workforce",
    "out of workforce", "retired"
}

SELF_EMPLOYED_BUCKET = {
    "self employed", "self-employed", "selfemployed", "self",
    "freelance", "freelancer", "independent", "independent contractor",
    "contract", "own business", "own company", "private"
}

def normalize_employer(emp):
    """
    Normalizes an employer name for comparison:
    - Strips whitespace, lowercases
    - Maps no-employer variations to a single token
    - Maps self-employed variations to a single token
    - Otherwise returns the lowercased string
    """
    if not emp:
        return "__no_employer__"
    cleaned = emp.lower().strip()
    if cleaned in NO_EMPLOYER_BUCKET:
        return "__no_employer__"
    if cleaned in HOMEMAKER_EQUIVALENTS:
        return "__no_employer__"
    if cleaned in SELF_EMPLOYED_BUCKET:
        return "__self_employed__"
    # Remove common suffixes that don't change identity
    for suffix in [" llc", " inc", " corp", " co", " ltd", " lp", " llp",
                   ", llc", ", inc", ", corp", ", co", ", ltd", " & co",
                   " and co", " corporation", " company"]:
        if cleaned.endswith(suffix):
            cleaned = cleaned[:-len(suffix)].strip()
    return cleaned


def employer_partial_match(emp1, emp2):
    """
    Returns True if two employer names are considered a partial match.
    Uses two-word match if both names have 2+ words, otherwise single-word match.
    Both comparisons are case-insensitive.
    """
    if not emp1 or not emp2:
        return False
    words1 = emp1.lower().split()
    words2 = emp2.lower().split()
    if len(words1) >= 2 and len(words2) >= 2:
        # Two-word match: both first two words must appear in the other name
        return words1[0] in words2 and words1[1] in words2
    else:
        # Single-word fallback
        return words1[0] in words2 or words2[0] in words1


def employers_are_same(emp1, emp2):
    """
    Returns True if two employer strings should be treated as the same employer.
    Uses normalization first, then falls back to partial match.
    """
    n1 = normalize_employer(emp1)
    n2 = normalize_employer(emp2)
    # Both normalized to same token (including __no_employer__ and __self_employed__)
    if n1 == n2:
        return True
    # Partial two-word match on normalized names
    return employer_partial_match(n1, n2)


# Pre-compiled regex patterns for street normalization (faster than re.sub each call)
_RE_PUNCT        = re.compile(r'[^a-z0-9 ]')
_RE_MULTI_SPACE  = re.compile(r' +')
_STREET_ABBREVS  = [
    (re.compile(r'\bstreet\b'),    'st'),
    (re.compile(r'\bavenue\b'),    'ave'),
    (re.compile(r'\bboulevard\b'), 'blvd'),
    (re.compile(r'\bdrive\b'),     'dr'),
    (re.compile(r'\broad\b'),      'rd'),
    (re.compile(r'\blane\b'),      'ln'),
    (re.compile(r'\bcourt\b'),     'ct'),
    (re.compile(r'\bplace\b'),     'pl'),
    (re.compile(r'\bcircle\b'),    'cir'),
    (re.compile(r'\bterrace\b'),   'ter'),
    (re.compile(r'\bnorth\b'),     'n'),
    (re.compile(r'\bsouth\b'),     's'),
    (re.compile(r'\beast\b'),      'e'),
    (re.compile(r'\bwest\b'),      'w'),
    (re.compile(r'\bapartment\b'), 'apt'),
    (re.compile(r'\bsuite\b'),     'ste'),
    (re.compile(r'\bfloor\b'),     'fl'),
    (re.compile(r'\bbuilding\b'),  'bldg'),
]


def normalize_street(addr):
    """
    Normalizes a street address for comparison.
    Lowercases, strips punctuation, standardizes common abbreviations.
    Handles both spelled-out and abbreviated forms (e.g. STREET vs ST).
    Uses pre-compiled regex patterns for speed.
    """
    if not addr:
        return ""
    addr = addr.lower().strip()
    addr = _RE_PUNCT.sub(' ', addr)
    for pattern, replacement in _STREET_ABBREVS:
        addr = pattern.sub(replacement, addr)
    # Remove unit/apt/suite/floor suffixes entirely for core address comparison
    addr = re.sub(r'\b(apt|ste|suite|unit|fl|floor|bldg|#)\b.*$', '', addr).strip()
    return _RE_MULTI_SPACE.sub(' ', addr).strip()


def streets_match(addr1, addr2):
    """
    Returns True if two street addresses are considered the same.
    Compares street number + first word of street name to handle
    apt/suite differences gracefully.
    """
    n1 = normalize_street(addr1)
    n2 = normalize_street(addr2)
    if not n1 or not n2:
        return False
    if n1 == n2:
        return True
    def key_tokens(s):
        tokens = s.split()
        return tokens[:2] if len(tokens) >= 2 else tokens
    return key_tokens(n1) == key_tokens(n2)


def build_bulk_index(donors):
    """
    Downloads FEC bulk CSV files (one per election cycle) and builds an
    in-memory index keyed by (last_name_lower, first_name_lower, city_lower, state_lower).

    Only downloads years relevant to donors in the list.
    Skips download if a cached file already exists locally.
    Called ONCE at startup before parallel processing begins.
    """
    import urllib.request
    import zipfile
    import io
    import csv
    from collections import defaultdict

    global _BULK_INDEX
    os.makedirs(BULK_CACHE_FOLDER, exist_ok=True)

    # Build a set of (last, first, city, state) we actually need to look up
    donor_keys = set()
    for d in donors:
        donor_keys.add((
            d["last_name"].lower().strip(),
            d["first_name"].lower().strip(),
            d["city"].lower().strip(),
            d["state"].lower().strip(),
        ))

    years = list(range(2016, 2027, 2))  # FEC even-year cycles
    index = defaultdict(list)

    for year in years:
        zip_filename = os.path.join(BULK_CACHE_FOLDER, f"indiv{str(year)[2:]}.zip")
        url = f"https://www.fec.gov/files/bulk-downloads/{year}/indiv{str(year)[2:]}.zip"

        # Download only if not already cached
        if not os.path.exists(zip_filename):
            print(f"  Downloading FEC bulk data for {year} cycle (this may take a few minutes)...")
            try:
                urllib.request.urlretrieve(url, zip_filename)
                print(f"  Downloaded {year} cycle.")
            except Exception as e:
                print(f"  WARNING — Could not download {year} cycle: {e}")
                continue
        else:
            print(f"  Using cached FEC bulk data for {year} cycle.")

        # Parse and index only records matching our donor list
        try:
            with zipfile.ZipFile(zip_filename) as z:
                txt_name = [n for n in z.namelist() if n.startswith("itcont")][0]
                with z.open(txt_name) as f:
                    reader = csv.reader(io.TextIOWrapper(f, encoding="latin-1"), delimiter="|")
                    matched = 0
                    for row in reader:
                        # FEC itcont pipe-delimited layout:
                        # 7:NAME  8:CITY  9:STATE  10:ZIP  11:EMPLOYER
                        # 12:OCCUPATION  13:DATE  14:AMOUNT
                        if len(row) < 15:
                            continue
                        r_city  = row[8].lower().strip()
                        r_state = row[9].lower().strip()
                        name_raw = row[7].strip()  # "LAST, FIRST MIDDLE"

                        # Parse name — FEC format is "LAST, FIRST" or "LAST, FIRST MIDDLE"
                        if "," in name_raw:
                            parts = name_raw.split(",", 1)
                            r_last  = parts[0].lower().strip()
                            r_first = parts[1].lower().strip().split()[0] if parts[1].strip() else ""
                        else:
                            continue  # skip malformed

                        key = (r_last, r_first, r_city, r_state)
                        if key not in donor_keys:
                            continue

                        index[key].append({
                            "name":       name_raw,
                            "city":       row[8].strip(),
                            "state":      row[9].strip(),
                            "zip":        row[10].strip(),
                            "employer":   row[11].strip(),
                            "occupation": row[12].strip(),
                            "date":       row[13].strip(),
                            "amount":     row[14].strip(),
                            "year_cycle": year,
                        })
                        matched += 1

            print(f"  Indexed {matched} matching records from {year} cycle.")
        except Exception as e:
            print(f"  WARNING — Could not parse {year} bulk file: {e}")

    _BULK_INDEX = dict(index)
    total_indexed = sum(len(v) for v in _BULK_INDEX.values())
    print(f"\nBulk index built: {total_indexed} records across {len(_BULK_INDEX)} donors.\n")


def check_bulk_index(last_name, first_name, city, state, donor_address):
    """
    Looks up a donor in the pre-built bulk index.
    Compares their ZIP from the spreadsheet address against FEC records.
    Returns (confirmed, flagged, notes).

    confirmed: records whose ZIP matches the donor address ZIP
    flagged:   records with a different ZIP (possible different person or old address)
    """
    key = (last_name.lower().strip(), first_name.lower().strip(),
           city.lower().strip(), state.lower().strip())

    records = _BULK_INDEX.get(key, [])
    if not records:
        return [], [], ["BULK CHECK — No records found in bulk index for this donor."]

    # Extract ZIP from donor address (last 5 digits of the address string, or standalone)
    zip_match = re.search(r'\b(\d{5})\b', donor_address)
    donor_zip = zip_match.group(1) if zip_match else ""

    confirmed = []
    flagged   = []
    notes     = []

    for r in records:
        r_zip = r.get("zip", "")[:5]
        if donor_zip and r_zip and r_zip == donor_zip:
            confirmed.append(r)
        elif donor_zip and r_zip and r_zip != donor_zip:
            flagged.append(r)
        else:
            # No ZIP to compare — include but note
            confirmed.append(r)

    if flagged:
        flagged_zips = list({r.get("zip", "?")[:5] for r in flagged})
        notes.append(
            f"BULK CHECK — {len(confirmed)} record(s) match donor ZIP ({donor_zip}); "
            f"{len(flagged)} record(s) found under different ZIP(s): {', '.join(flagged_zips)}. "
            f"Flagged records included in evaluation but may be a different person."
        )
    else:
        notes.append(
            f"BULK CHECK — {len(confirmed)} record(s) confirmed via ZIP match ({donor_zip})."
        )

    return confirmed, flagged, notes


def dedupe_actblue(contributions):
    """
    Removes ActBlue (or any configured fundraising-platform pass-through) duplicates.
    When a donor gives through a platform like ActBlue, FEC records it twice:
    once to the platform's own committee and once to the actual campaign committee.
    Keeps the real committee record and drops the pass-through one.
    If the pass-through is the only record for a given date+amount, keeps it.

    PLATFORM_KEYWORDS below covers ActBlue (Democratic-side) by default.
    If you're tracking Republican-side donors, add WinRed; for other
    platforms, add their committee name keywords here too.
    """
    PLATFORM_KEYWORDS = {"actblue", "act blue"}

    def is_passthrough(c):
        cmte_name = ((c.get("committee", {}) or {}).get("name") or "").strip().lower()
        return any(k in cmte_name for k in PLATFORM_KEYWORDS)

    from collections import defaultdict
    date_amount_groups = defaultdict(list)
    for c in contributions:
        key = (
            (c.get("contribution_receipt_date") or "")[:10],
            c.get("contribution_receipt_amount", 0)
        )
        date_amount_groups[key].append(c)

    deduped = []
    for key, group in date_amount_groups.items():
        real_recs        = [c for c in group if not is_passthrough(c)]
        passthrough_recs = [c for c in group if is_passthrough(c)]
        if real_recs:
            deduped.extend(real_recs)
        else:
            deduped.extend(passthrough_recs)

    return deduped


def fetch_contributions(last_name, first_name, city, state, company, address="", occupation=""):
    """
    Pulls contributions using a tiered search strategy:

    Step 1:  Last + First + State via API.
             - If employers all match → done (if >= 3 results).
             - If employers diverge → address/city/employer disambiguation.

    Step 1b: Single-letter first name only — search by last name + state,
             filter client-side to records whose first name starts with that letter.
             Runs if Step 1 returns 0 results.

    Step 2:  If < 3 results after Step 1/1b, try Last + First + Employer (partial).
             Flag results where city/state differ.

    Step 3:  If still < 3 results, flag for manual review.

    Platform pass-through deduplication (e.g. ActBlue) runs on all results
    before any return.

    Returns (contributions, search_notes).
    """
    search_notes = []

    base_params = {
        "contributor_name": f"{last_name}, {first_name}",
        "min_date": MIN_DATE,
        "per_page": 100,
        "sort": "-contribution_receipt_amount",
        "api_key": API_KEY,
    }

    # --- Step 1: First + Last + State (no city filter — city names vary in FEC records)
    # Address matching handles disambiguation instead of city filtering
    params_step1 = dict(base_params)
    if state:
        params_step1["contributor_state"] = state

    results = _fetch_raw(params_step1)

    # Only try flipped name format if first attempt returns zero results
    if not results:
        params_flip = dict(params_step1)
        params_flip["contributor_name"] = f"{first_name} {last_name}"
        results = _fetch_raw(params_flip)

    # Filter results to only those matching the donor's state
    # (since we dropped city, we rely on state + address matching)
    if results and state:
        results = [
            r for r in results
            if (r.get("contributor_state") or "").strip().upper() == state.upper()
        ]

    if results:
        # Collect unique employer names from results
        employers_in_results = list({
            (r.get("contributor_employer") or "").strip()
            for r in results
            if (r.get("contributor_employer") or "").strip()
        })

        # Check if all employers are the same (after normalization + partial match)
        all_match = True
        if len(employers_in_results) > 1:
            reference = employers_in_results[0]
            for emp in employers_in_results[1:]:
                if not employers_are_same(reference, emp):
                    all_match = False
                    break

        if not all_match:
            # Employers diverge — flag for manual review instead of bulk download

            if address:
                addr_confirmed  = []  # street address matches
                city_confirmed  = []  # street doesn't match but city matches
                addr_flagged    = []  # neither matches
                addr_unknown    = []  # no street in FEC record

                # Pre-build known employers from ALL results before the loop
                # so we don't miss records processed before confirmed ones are seen
                all_employers_in_results = set(
                    (r2.get("contributor_employer") or "").strip()
                    for r2 in results
                    if (r2.get("contributor_employer") or "").strip()
                )
                # Also check if all/most results share the same no-employer status
                no_employer_count = sum(
                    1 for r2 in results
                    if normalize_employer((r2.get("contributor_employer") or "").strip()) == "__no_employer__"
                )
                majority_no_employer = no_employer_count > len(results) / 2

                for r in results:
                    r_street = (r.get("contributor_street_1") or "").strip()
                    r_city   = (r.get("contributor_city") or "").strip().lower()
                    r_state  = (r.get("contributor_state") or "").strip().upper()

                    if not r_street:
                        addr_unknown.append(r)
                    elif streets_match(address, r_street):
                        addr_confirmed.append(r)
                    elif r_city == city.lower():
                        # Street differs but city matches — likely same person, old address
                        city_confirmed.append(r)
                    else:
                        r_emp = (r.get("contributor_employer") or "").strip()
                        # Neither street nor city matches — check employer as last resort.
                        # Use pre-built employer set from all results to avoid ordering issues.
                        emp_match = (
                            (company and employers_are_same(r_emp, company))
                            or any(employers_are_same(r_emp, ke) for ke in all_employers_in_results)
                            or (normalize_employer(r_emp) == "__no_employer__"
                                and normalize_employer(company) == "__no_employer__")
                            or (normalize_employer(r_emp) == "__no_employer__"
                                and majority_no_employer)
                            or (r_state == state.upper()
                                and normalize_employer(r_emp) == "__no_employer__"
                                and normalize_employer(company) == "__no_employer__")
                        )
                        if emp_match:
                            city_confirmed.append(r)  # treat as same person, flag it
                        else:
                            addr_flagged.append(r)

                if addr_flagged or city_confirmed:
                    flagged_streets = list({
                        (r.get("contributor_street_1") or "?").strip()
                        for r in addr_flagged
                    })
                    city_streets = list({
                        (r.get("contributor_street_1") or "?").strip()
                        for r in city_confirmed
                    })

                    note_parts = [
                        f"NOTE — Employer mismatch detected. Address verification: "
                        f"{len(addr_confirmed)} record(s) match street address ({address})"
                    ]
                    city_only = [r for r in city_confirmed
                                  if (r.get("contributor_city") or "").strip().lower() == city.lower()]
                    emp_only  = [r for r in city_confirmed
                                  if (r.get("contributor_city") or "").strip().lower() != city.lower()]

                    if city_only:
                        note_parts.append(
                            f"; {len(city_only)} record(s) matched city only "
                            f"(different street: {', '.join(city_streets[:3])}) — "
                            f"included as likely same person (possible old address)"
                        )
                    if emp_only:
                        note_parts.append(
                            f"; {len(emp_only)} record(s) matched employer only "
                            f"(different city/address) — "
                            f"included as likely same person (possible relocation)"
                        )
                    if addr_flagged:
                        note_parts.append(
                            f"; {len(addr_flagged)} record(s) found at different "
                            f"city/address: {', '.join(flagged_streets[:3])} — "
                            f"excluded, may be a different person"
                        )
                    search_notes.append("".join(note_parts) + ".")

                    # Priority order: street match > employer match > city match (last resort)
                    # Split city_confirmed into true city-only vs employer-matched
                    city_only_recs = [r for r in city_confirmed
                                      if (r.get("contributor_city") or "").strip().lower() == city.lower()]
                    emp_matched_recs = [r for r in city_confirmed
                                        if (r.get("contributor_city") or "").strip().lower() != city.lower()]

                    # Always include street-confirmed + employer-matched + unknowns
                    street_and_emp = addr_confirmed + emp_matched_recs + addr_unknown

                    # Check if street + employer records alone meet consideration criteria
                    street_emp_max = max(
                        (r.get("contribution_receipt_amount", 0) for r in street_and_emp),
                        default=0
                    )

                    # From city-only records, always keep those where occupation matches
                    # spreadsheet occupation — consistent occupation across addresses is a
                    # strong signal it's the same person (e.g. donor moved but same job title)
                    _sheet_occ = occupation  # capture in local scope for lambda
                    city_occ_match = [
                        r for r in city_only_recs
                        if (r.get("contributor_occupation") or "").strip().lower() == _sheet_occ.strip().lower()
                        and _sheet_occ.strip()
                    ]
                    city_no_occ_match = [
                        r for r in city_only_recs
                        if r not in city_occ_match
                    ]

                    # Only add city-only records without occupation match as last resort
                    if len(street_and_emp) < 3 or street_emp_max < MIN_SINGLE_DONATION:
                        results = street_and_emp + city_occ_match + city_no_occ_match
                        if city_only_recs:
                            search_notes.append(
                                f"NOTE — City-matched records included as last resort "
                                f"({len(city_only_recs)} record(s)) since street+employer "
                                f"results were insufficient."
                            )
                    else:
                        # Always include occupation-matched city records
                        results = street_and_emp + city_occ_match
                        if city_occ_match:
                            search_notes.append(
                                f"NOTE — {len(city_occ_match)} city-matched record(s) included "
                                f"due to matching occupation ('{occupation}')."
                            )
                        if city_no_occ_match:
                            search_notes.append(
                                f"NOTE — {len(city_no_occ_match)} city-matched record(s) excluded "
                                f"(street+employer matches sufficient, occupation did not match)."
                            )

                    if not results:
                        results = street_and_emp + city_occ_match + city_no_occ_match + addr_flagged
                        search_notes.append(
                            f"WARNING — No confirmed records found for "
                            f"{first_name} {last_name}. Using all records — verify manually."
                        )
                else:
                    search_notes.append(
                        f"NOTE — Employer mismatch detected but all {len(results)} "
                        f"record(s) confirmed via street address match. Treated as same person."
                    )
            else:
                search_notes.append(
                    f"NOTE — Employer mismatch detected for {first_name} {last_name} "
                    f"({', '.join(employers_in_results[:5])}). No address in spreadsheet "
                    f"to verify — all records included. Recommend manual check."
                )
        else:
            # Employers all match — check if any address differs
            if len(employers_in_results) == 1 and company:
                fec_employer = employers_in_results[0]
                if fec_employer.lower() != company.lower() and employer_partial_match(fec_employer, company):
                    search_notes.append(
                        f'NOTE — Employer in FEC records ("{fec_employer}") partially matches '
                        f'your spreadsheet ("{company}"). Treated as same person.'
                    )

        if len(results) >= 3:
            results = dedupe_actblue(results)
            return results, search_notes

    # --- Step 1b: Single-letter first name — search by last name + state ---
    # The FEC API can't reliably match "McAllister, C" so we search by last name
    # only and filter client-side to records whose first name starts with that letter.
    if len(first_name.strip()) == 1 and not results:
        letter = first_name.strip().upper()
        params_1b = {
            "contributor_last_name": last_name,
            "min_date": MIN_DATE,
            "per_page": 100,
            "sort": "-contribution_receipt_amount",
            "api_key": API_KEY,
        }
        if state:
            params_1b["contributor_state"] = state

        lastname_results = _fetch_raw(params_1b)

        # Filter to records whose first name starts with the given letter
        def fec_first_initial(r):
            raw = (r.get("contributor_name") or "").strip()
            # FEC name format: "LAST, FIRST" or "LAST, FIRST MIDDLE"
            if "," in raw:
                first_part = raw.split(",", 1)[1].strip()
                return first_part[:1].upper() if first_part else ""
            return ""

        matched_1b = [r for r in lastname_results if fec_first_initial(r) == letter]

        if matched_1b:
            # Merge with any existing results (shouldn't be any, but be safe)
            existing_ids = {r.get("transaction_id") for r in results}
            new_1b = [r for r in matched_1b if r.get("transaction_id") not in existing_ids]
            results.extend(new_1b)
            search_notes.append(
                f"NOTE — Single-letter first name ('{first_name}'): searched by last name "
                f"'{last_name}' + state, then filtered to records with first initial '{letter}'. "
                f"Found {len(new_1b)} matching record(s)."
            )
        else:
            search_notes.append(
                f"NOTE — Single-letter first name ('{first_name}'): last-name search returned "
                f"no records with first initial '{letter}' in {state}."
            )

        if len(results) >= 3:
            results = dedupe_actblue(results)
            return results, search_notes

    # --- Step 2: First + Last + Employer (partial, two-word) ---
    # Only runs if Step 1 returned fewer than 3 results
    if company and normalize_employer(company) not in ("__no_employer__", "__self_employed__", "", "nan"):
        words = company.split()
        if len(words) >= 2:
            company_keyword = f"{words[0]} {words[1]}"
        else:
            company_keyword = words[0] if words else company

        params_step2 = dict(base_params)
        params_step2["contributor_employer"] = company_keyword

        employer_results = _fetch_raw(params_step2)
        if not employer_results:
            params_step2["contributor_name"] = f"{first_name} {last_name}"
            employer_results = _fetch_raw(params_step2)

        if employer_results:
            mismatched = []
            matched    = []
            for r in employer_results:
                r_city  = (r.get("contributor_city")  or "").strip().lower()
                r_state = (r.get("contributor_state") or "").strip().lower()
                r_emp   = (r.get("contributor_employer") or "").strip()

                # Same employer (normalized) → flag address diff but treat as same person
                if employers_are_same(r_emp, company) and r_state != state.lower():
                    mismatched.append(r)
                    search_notes.append(
                        f'VERIFY — Contribution found with exact employer match ("{r_emp}") '
                        f"but different location: {r_city.title()}, {r_state.upper()}. "
                        f"Treated as same person but recommend confirming."
                    )
                elif r_city != city.lower() or r_state != state.lower():
                    mismatched.append(r)
                else:
                    matched.append(r)

            combined = matched + mismatched
            existing_ids = {r.get("transaction_id") for r in results}
            new_results  = [r for r in combined if r.get("transaction_id") not in existing_ids]
            results.extend(new_results)

            if mismatched and not any("exact employer match" in n for n in search_notes):
                mismatch_locations = list({
                    f"{(r.get('contributor_city') or '').title()}, {(r.get('contributor_state') or '').upper()}"
                    for r in mismatched
                })
                search_notes.append(
                    f"VERIFY — {len(mismatched)} contribution(s) found via employer search "
                    f'("{company}") under a different city/state: '
                    f"{', '.join(mismatch_locations)}. Included in evaluation but may be "
                    f"a different person. Recommend manual FEC check."
                )
            if new_results:
                search_notes.append(
                    f"NOTE — Step 1 (city/state search) returned fewer than 3 results. "
                    f"{len(new_results)} additional contribution(s) found via employer "
                    f"search and merged into this record."
                )

    # --- Deduplicate platform pass-throughs (e.g. ActBlue) before returning ---
    results = dedupe_actblue(results)

    # --- Step 3: Flag if still thin ---
    if len(results) < 3:
        search_notes.append(
            f"MANUAL REVIEW RECOMMENDED — Only {len(results)} contribution(s) found after "
            f"searching by name+city+state and name+employer. Donor may have contributions "
            f"under a previous address or state not captured here. "
            f"Check manually at: https://www.fec.gov/data/receipts/individual-contributions/"
            f"?contributor_name={first_name}+{last_name}"
        )

    return results, search_notes


def format_amount(amount):
    """Formats a number as $X,XXX"""
    try:
        return f"${int(float(amount)):,}"
    except:
        return "$0"


def format_date(date_str):
    """Converts YYYY-MM-DD to MM-DD-YYYY"""
    try:
        dt = datetime.strptime(str(date_str)[:10], "%Y-%m-%d")
        return dt.strftime("%m-%d-%Y")
    except:
        return str(date_str)


def parse_date(date_str):
    """Returns a date object from a YYYY-MM-DD string, or None."""
    try:
        return datetime.strptime(str(date_str)[:10], "%Y-%m-%d").date()
    except:
        return None


def is_disqualifying_employer(company):
    """Returns the matched employer name if disqualifying, else None.

    Uses whole-word matching so short entries (e.g. "td") only match
    standalone words/phrases and not substrings inside other words
    (e.g. "td" won't match "Ltd"). Periods are stripped and extra
    whitespace collapsed before matching, so punctuated variants like
    "J.P. Morgan" or "J. P. Morgan" are still caught by "jp morgan".

    Edit DISQUALIFYING_EMPLOYERS in config.json — not this function —
    to change which employers trigger this.
    """
    company_lower = company.lower().replace(".", "")
    company_lower = re.sub(r"\s+", " ", company_lower).strip()
    for employer in DISQUALIFYING_EMPLOYERS:
        pattern = r"\b" + re.escape(employer) + r"\b"
        if re.search(pattern, company_lower):
            return employer
    return None


def is_disqualifying_occupation(state, occupation):
    """Returns True if the occupation is disqualifying for this donor's state.
    Looks up DISQUALIFYING_OCCUPATIONS_BY_STATE[state] from config.json.
    """
    occ_lower = occupation.lower()
    for occ in DISQUALIFYING_OCCUPATIONS_BY_STATE.get(state.upper(), []):
        if occ in occ_lower:
            return True
    return False


def needs_manual_verification(state, occupation):
    """Returns True if this occupation should be flagged for manual review
    in this donor's state, per VERIFY_OCCUPATIONS_BY_STATE in config.json.
    """
    return occupation.lower().strip() in VERIFY_OCCUPATIONS_BY_STATE.get(state.upper(), [])


# =========================================================
# EVALUATE_DONOR — THIS IS THE FUNCTION YOU WILL EDIT
# =========================================================
# Everything above this point is generic "engine" code (talking to the FEC
# API, matching addresses/employers, deduping records) that most users
# won't need to touch.
#
# This function is different: it's where your actual eligibility judgment
# lives. The example below is a GENERIC PLACEHOLDER that demonstrates the
# *shape* most real-world criteria take:
#
#   1. Hard disqualifiers (employer, occupation) — checked first, short-circuit
#   2. A "no data" floor — donors with no giving history are out
#   3. A minimum single-donation floor — too-small donors are out
#   4. A "special state" path — one specific state/region with its own
#      tiered rules (delete this whole block if every donor should be
#      treated identically — see config.json: set special_state.code to "")
#   5. A "standard path" for everyone else — also tiered by dollar amount
#      and recency
#
# Real eligibility criteria are rarely a single number — they're usually a
# few "yes, but only if..." branches like this. Don't try to flatten your
# rules into a single config number if they don't actually work that way;
# write the branches out here instead. See PROMPTS.md for example prompts
# to hand an AI assistant when adapting this to your own rules.
# =========================================================

def evaluate_donor(contributions, state, company, occupation):
    """
    Evaluates whether a donor should be considered for outreach.

    PLACEHOLDER rules (in order) — replace with your own:
    1. Missing name → skipped before this function is called
    2. Disqualifying employer (config.json) → DO NOT CONSIDER
    3. Disqualifying occupation for this state (config.json) → DO NOT CONSIDER
    4. No contributions at all → DO NOT CONSIDER
    5. No single donation over MIN_SINGLE_DONATION → DO NOT CONSIDER
    6. Special-state donors: tiered by donation count/recency (config.json)
    7. Everyone else ("standard path"): tiered by dollar amount + recency

    Returns a detailed reasoning string explaining the call.
    """
    today = date.today()
    standard_recency_cutoff = date(today.year - STANDARD_RECENCY_YEARS, today.month, today.day)

    # ---------------------------------------------------------
    # STEP 2 — Hard disqualifier: employer
    # Edit: config.json -> eligibility.disqualifying_employers
    # ---------------------------------------------------------
    matched_employer = is_disqualifying_employer(company)
    if matched_employer:
        return (
            f"DO NOT CONSIDER — Donor's employer ({company}) is on the "
            f"disqualifying employer list. Excluded regardless of giving history."
        )

    # ---------------------------------------------------------
    # STEP 3 — Hard disqualifier: occupation, by state
    # Edit: config.json -> eligibility.disqualifying_occupations_by_state
    # ---------------------------------------------------------
    if is_disqualifying_occupation(state, occupation):
        return (
            f"DO NOT CONSIDER — Donor's occupation ({occupation}) is on the "
            f"disqualifying occupation list for {state.upper()}."
        )

    # Flag (but don't disqualify) occupations that need a human to double-check.
    # Edit: config.json -> eligibility.verify_occupations_by_state
    manual_review_flag = needs_manual_verification(state, occupation)
    review_note = (
        f" ⚠ MANUAL REVIEW — Occupation listed as '{occupation}': verify before contacting."
        if manual_review_flag else ""
    )

    # ---------------------------------------------------------
    # STEP 4 — No giving history at all
    # ---------------------------------------------------------
    if not contributions:
        return f"DO NOT CONSIDER — No federal giving history found since {MIN_DATE}."

    amounts       = [c.get("contribution_receipt_amount", 0) for c in contributions]
    dates         = [parse_date(c.get("contribution_receipt_date", "")) for c in contributions]
    max_donation  = max(amounts)
    num_donations = len(contributions)

    # ---------------------------------------------------------
    # STEP 5 — Minimum single-donation floor
    # Edit: config.json -> eligibility.min_single_donation
    # ---------------------------------------------------------
    if max_donation < MIN_SINGLE_DONATION:
        return (
            f"DO NOT CONSIDER — Highest single donation is {format_amount(max_donation)}, "
            f"below the {format_amount(MIN_SINGLE_DONATION)} minimum threshold."
        )

    # ---------------------------------------------------------
    # STEP 6 — SPECIAL STATE PATH
    # Set config.json -> eligibility.special_state.code to "" to skip this
    # block entirely and send every donor down the standard path below.
    #
    # Example shape: a region gets evaluated on *how many times* they've
    # given recently, with a secondary dollar-amount tier for donors who
    # don't meet the count but have given generously at least once.
    # ---------------------------------------------------------
    if SPECIAL_STATE_CODE and state.upper() == SPECIAL_STATE_CODE:
        special_cutoff = date(today.year - SPECIAL_HIGH_TIER_YEARS, today.month, today.day)
        recent_donations = [d for d in dates if d and d >= special_cutoff]
        num_recent = len(recent_donations)

        over_mid_tier = [a for a in amounts if a >= SPECIAL_MID_TIER_AMOUNT]

        if num_recent >= SPECIAL_HIGH_TIER_COUNT:
            return (
                f"CONSIDER (high tier) — {SPECIAL_STATE_CODE} donor with {num_recent} "
                f"donation(s) in the last {SPECIAL_HIGH_TIER_YEARS} years "
                f"(meets the {SPECIAL_HIGH_TIER_COUNT}-donation threshold). "
                f"Highest single gift: {format_amount(max_donation)}. "
                f"Total donations on record since {MIN_DATE}: {num_donations}."
                + review_note
            )
        elif over_mid_tier:
            return (
                f"CONSIDER (mid tier) — {SPECIAL_STATE_CODE} donor with {num_recent} "
                f"donation(s) in the last {SPECIAL_HIGH_TIER_YEARS} years (below the "
                f"{SPECIAL_HIGH_TIER_COUNT}-donation threshold), but has "
                f"{len(over_mid_tier)} gift(s) of {format_amount(SPECIAL_MID_TIER_AMOUNT)} or more. "
                f"Highest single gift: {format_amount(max_donation)}. "
                f"Total donations on record since {MIN_DATE}: {num_donations}."
                + review_note
            )
        else:
            return (
                f"DO NOT CONSIDER — {SPECIAL_STATE_CODE} donor does not meet the repeated-giving "
                f"threshold ({num_recent} donation(s) in last {SPECIAL_HIGH_TIER_YEARS} years, "
                f"minimum is {SPECIAL_HIGH_TIER_COUNT}) and has no gifts of "
                f"{format_amount(SPECIAL_MID_TIER_AMOUNT)} or more. "
                f"Highest gift: {format_amount(max_donation)}."
            )

    # ---------------------------------------------------------
    # STEP 7 — STANDARD PATH (everyone not covered by the special-state block)
    # Edit: config.json -> eligibility.standard_path
    #
    # Example shape: tiered by dollar amount and recency. A very large
    # gift (high_value_threshold) qualifies on its own. A smaller-but-still
    # substantial gift (mid_value_threshold) needs to show up at least
    # mid_value_min_count times within recency_years to count as a pattern
    # rather than a one-off.
    # ---------------------------------------------------------
    over_mid = [(a, d) for a, d in zip(amounts, dates) if a >= STANDARD_MID_VALUE_THRESHOLD]
    recent_over_mid = [(a, d) for a, d in over_mid if d and d >= standard_recency_cutoff]

    if not recent_over_mid:
        over_high_any_time = [(a, d) for a, d in over_mid if a >= STANDARD_HIGH_VALUE_THRESHOLD]
        if over_high_any_time:
            return (
                f"CONSIDER — Donor ({state}) has no donations of "
                f"{format_amount(STANDARD_MID_VALUE_THRESHOLD)}+ within the last "
                f"{STANDARD_RECENCY_YEARS} years, but has {len(over_high_any_time)} "
                f"donation(s) of {format_amount(STANDARD_HIGH_VALUE_THRESHOLD)} or more on record. "
                f"High-value giving history overrides the recency requirement. "
                f"Highest gift: {format_amount(max_donation)}."
                + review_note
            )
        else:
            return (
                f"DO NOT CONSIDER — Donor ({state}) has no donations of "
                f"{format_amount(STANDARD_MID_VALUE_THRESHOLD)}+ within the last "
                f"{STANDARD_RECENCY_YEARS} years and no donations of "
                f"{format_amount(STANDARD_HIGH_VALUE_THRESHOLD)} or more on record. "
                f"Highest gift: {format_amount(max_donation)}."
            )

    recent_high = [(a, d) for a, d in recent_over_mid if a >= STANDARD_HIGH_VALUE_THRESHOLD]
    if recent_high:
        return (
            f"CONSIDER — Donor ({state}) has {len(recent_high)} donation(s) of "
            f"{format_amount(STANDARD_HIGH_VALUE_THRESHOLD)} or more within the last "
            f"{STANDARD_RECENCY_YEARS} years. Highest single gift: {format_amount(max_donation)}. "
            f"Total donations on record since {MIN_DATE}: {num_donations}."
            + review_note
        )

    if len(recent_over_mid) < STANDARD_MID_VALUE_MIN_COUNT:
        return (
            f"CONSIDER (medium quality) — Donor ({state}) has {len(recent_over_mid)} "
            f"recent donation(s) of {format_amount(STANDARD_MID_VALUE_THRESHOLD)}+ within the "
            f"last {STANDARD_RECENCY_YEARS} years, below the preferred count of "
            f"{STANDARD_MID_VALUE_MIN_COUNT}. Giving history is present but limited. "
            f"Highest gift: {format_amount(max_donation)}. "
            f"Total donations on record since {MIN_DATE}: {num_donations}."
            + review_note
        )

    return (
        f"CONSIDER — Donor ({state}) meets threshold: {len(recent_over_mid)} donation(s) of "
        f"{format_amount(STANDARD_MID_VALUE_THRESHOLD)}+ within the last {STANDARD_RECENCY_YEARS} "
        f"years (minimum is {STANDARD_MID_VALUE_MIN_COUNT}). Highest single gift: "
        f"{format_amount(max_donation)}. Total donations on record since {MIN_DATE}: {num_donations}."
        + review_note
    )


def save_donor_file(donor, contributions, call, search_notes, folder):
    """Saves an individual .txt file for one donor."""
    first      = donor["first_name"]
    last       = donor["last_name"]
    state      = donor["state"]
    city       = donor["city"]
    company    = donor["company"]
    occupation = donor["occupation"]

    filename = f"fec_{last}_{first}.txt".replace(" ", "_")
    filepath = os.path.join(folder, filename)

    with open(filepath, "w") as f:
        f.write(f"Federal Giving History: {first} {last}\n")
        f.write(f"Location: {city}, {state}\n")
        if company and company.lower() != "nan":
            f.write(f"Employer: {company}\n")
        if occupation and occupation.lower() != "nan":
            f.write(f"Occupation: {occupation}\n")
        f.write("\n")

        if search_notes:
            f.write("SEARCH NOTES:\n")
            for note in search_notes:
                f.write(f"  - {note}\n")
            f.write("\n")

        if contributions:
            f.write("Federal Giving History:\n")
            for c in contributions:
                date_str  = format_date(c.get("contribution_receipt_date", ""))
                amount    = format_amount(c.get("contribution_receipt_amount", 0))
                committee = ((c.get("committee") or {}).get("name") or "Unknown Committee").title()
                f.write(f"{date_str} // {amount} // {committee}\n")
        else:
            f.write(f"No federal contributions found since {MIN_DATE}.\n")

        f.write(f"\nCONSIDERATION CALL:\n{call}\n")

    return filepath


def main():
    # --- Validate setup ---
    if not API_KEY or API_KEY == "YOUR_FEC_API_KEY_HERE":
        print("ERROR: You need to add your FEC API key.")
        print("Get one free at: https://api.data.gov/signup/")
        print("Paste it into the \"api_key\" field in config.json.")
        return

    if not os.path.exists(INPUT_FILE):
        print(f"ERROR: Could not find input file: {INPUT_FILE}")
        print("Check the \"input_file\" path in config.json.")
        return

    # --- Load donors ---
    donors = load_donor_list(INPUT_FILE)

    # --- Create output folder ---
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # --- Process each donor in parallel ---
    summary_rows = []
    total = len(donors)

    def process_donor(args):
        """Processes a single donor — designed to run in a thread."""
        i, donor = args
        first      = donor["first_name"]
        last       = donor["last_name"]
        city       = donor["city"]
        state      = donor["state"]
        company    = donor["company"]
        occupation = donor["occupation"]
        address    = donor.get("address", "")

        print(f"\n[{i}/{total}] Starting: {first} {last}, {city}, {state}", flush=True)
        start_time = time.time()

        contributions, search_notes = fetch_contributions(last, first, city, state, company, address, occupation)
        contributions.sort(key=lambda x: x.get("contribution_receipt_amount", 0), reverse=True)

        # Filter to top results over the minimum single-donation threshold —
        # matches what's shown in the txt file
        contributions = [
            c for c in contributions
            if c.get("contribution_receipt_amount", 0) >= MIN_SINGLE_DONATION
        ][:MAX_RESULTS_PER_DONOR]

        call = evaluate_donor(contributions, state, company, occupation)

        elapsed = round(time.time() - start_time, 1)
        print(f"[{i}/{total}] Done ({elapsed}s): {first} {last} → {call[:80]}...", flush=True)
        for note in search_notes:
            print(f"    ⚠ {note}", flush=True)

        save_donor_file(donor, contributions, call, search_notes, OUTPUT_FOLDER)

        max_gift = max((c.get("contribution_receipt_amount", 0) for c in contributions), default=0)
        return {
            "Last Name":        last,
            "First Name":       first,
            "City":             city,
            "State":            state,
            "Company":          company,
            "Occupation":       occupation,
            "Total Donations":  len(contributions),
            "Highest Gift":     format_amount(max_gift),
            "Consideration":    call,
        }

    print(f"\nProcessing {total} donors with {MAX_WORKERS} parallel workers...\n", flush=True)

    DONOR_TIMEOUT = 90  # seconds before giving up on a single donor

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(process_donor, (i, donor)): (i, donor)
                   for i, donor in enumerate(donors, start=1)}
        for future in as_completed(futures, timeout=None):
            i, donor = futures[future]
            first = donor["first_name"]
            last  = donor["last_name"]
            try:
                row = future.result(timeout=DONOR_TIMEOUT)
                summary_rows.append(row)
            except TimeoutError:
                print(f"[{i}] TIMEOUT ({DONOR_TIMEOUT}s) — {first} {last} took too long, skipping. Check manually.", flush=True)
                summary_rows.append({
                    "Last Name":       last,
                    "First Name":      first,
                    "City":            donor.get("city", ""),
                    "State":           donor.get("state", ""),
                    "Company":         donor.get("company", ""),
                    "Occupation":      donor.get("occupation", ""),
                    "Total Donations": "N/A",
                    "Highest Gift":    "N/A",
                    "Consideration":   "TIMED OUT — check manually at fec.gov",
                })
            except Exception as e:
                print(f"[{i}] ERROR — {first} {last}: {e}", flush=True)
                summary_rows.append({
                    "Last Name":       last,
                    "First Name":      first,
                    "City":            donor.get("city", ""),
                    "State":           donor.get("state", ""),
                    "Company":         donor.get("company", ""),
                    "Occupation":      donor.get("occupation", ""),
                    "Total Donations": "N/A",
                    "Highest Gift":    "N/A",
                    "Consideration":   f"ERROR — {e}",
                })

    # Sort summary by original order (Last Name, First Name)
    summary_rows.sort(key=lambda r: (r["Last Name"], r["First Name"]))

    # --- Save summary CSV ---
    pd.DataFrame(summary_rows).to_csv(SUMMARY_CSV, index=False)

    print("\n" + "=" * 60)
    print(f"Done! Processed {len(donors)} donors.")
    print(f"Summary saved to:          {SUMMARY_CSV}")
    print(f"Individual files saved to: {OUTPUT_FOLDER}/")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print("FATAL ERROR:", e, flush=True)
        traceback.print_exc()
