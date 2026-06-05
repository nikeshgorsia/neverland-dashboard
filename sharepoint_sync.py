import os
import re
import requests
import pandas as pd
import numpy as np
from dotenv import load_dotenv
from msal import PublicClientApplication, SerializableTokenCache

load_dotenv()

def _get_secret(key):
    """Read from Streamlit secrets (cloud) or .env (local)."""
    try:
        import streamlit as st
        val = st.secrets.get(key)
        if val:
            return val
    except Exception:
        pass
    return os.getenv(key)

def _get_credentials():
    """Lazily load credentials so Streamlit secrets are available."""
    return {
        "TENANT_ID":     _get_secret("TENANT_ID"),
        "CLIENT_ID":     _get_secret("CLIENT_ID"),
        "CLIENT_SECRET": _get_secret("CLIENT_SECRET"),
        "FILE_URL":      _get_secret("SHAREPOINT_FILE_URL"),
        "BUDGET_URL":    _get_secret("BUDGET_FILE_URL"),
        "SCOPE_URL":     _get_secret("SCOPE_FILE_URL"),
    }

# Keep module-level vars for backward compat — loaded lazily on first use
TENANT_ID     = None
CLIENT_ID     = None
CLIENT_SECRET = None
FILE_URL      = None
BUDGET_URL    = None
SCOPE_URL     = None

def _load_globals():
    global TENANT_ID, CLIENT_ID, CLIENT_SECRET, FILE_URL, BUDGET_URL, SCOPE_URL
    creds = _get_credentials()
    TENANT_ID     = creds["TENANT_ID"]
    CLIENT_ID     = creds["CLIENT_ID"]
    CLIENT_SECRET = creds["CLIENT_SECRET"]
    FILE_URL      = creds["FILE_URL"]
    BUDGET_URL    = creds["BUDGET_URL"]
    SCOPE_URL     = creds["SCOPE_URL"]
SCOPES        = ["https://graph.microsoft.com/Files.Read",
                 "https://graph.microsoft.com/Sites.Read.All"]
CACHE_FILE    = os.path.expanduser("~/.neverland_token_cache.bin")

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

SECTIONS = ["CONFIRMED", "PROPOSED", "POTENTIAL", "ACCOUNT PLANNING", "SPECULATIVE"]


def _load_cache():
    cache = SerializableTokenCache()
    # 1. Try Streamlit secrets (cloud)
    try:
        import streamlit as st
        cached = st.secrets.get("TOKEN_CACHE") or st.session_state.get("_token_cache")
        if cached:
            cache.deserialize(cached)
            return cache
    except Exception:
        pass
    # 2. Fall back to local file
    if os.path.exists(CACHE_FILE):
        cache.deserialize(open(CACHE_FILE).read())
    return cache


def _save_cache(cache):
    if not cache.has_state_changed:
        return
    serialized = cache.serialize()
    # Save to local file
    try:
        open(CACHE_FILE, "w").write(serialized)
    except Exception:
        pass
    # Save to Streamlit session state for persistence within session
    try:
        import streamlit as st
        st.session_state["_token_cache"] = serialized
        st.session_state["_token_cache_str"] = serialized  # expose for secrets setup
    except Exception:
        pass


def get_device_flow():
    _load_globals()
    cache = _load_cache()
    app = PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return None, None, result["access_token"]

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise ValueError("Could not start login flow")
    return app, flow, None


def complete_device_flow(app, flow):
    result = app.acquire_token_by_device_flow(flow)
    _save_cache(app.token_cache)
    if "access_token" not in result:
        raise ValueError(f"Login failed: {result.get('error_description', result)}")
    return result["access_token"]


def _clean_value(val):
    if pd.isna(val):
        return 0.0
    s = str(val).replace("£", "").replace(",", "").replace(" ", "").strip()
    if s in ("", "-", "£0", "0"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _download_raw(token):
    _load_globals()
    headers = {"Authorization": f"Bearer {token}"}
    sharing_url = f"u!{__import__('base64').urlsafe_b64encode(FILE_URL.encode()).decode().rstrip('=')}"
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{sharing_url}/driveItem",
        headers=headers,
    )
    if resp.status_code == 200:
        item = resp.json()
        download_url = item.get("@microsoft.graph.downloadUrl") or \
            f"https://graph.microsoft.com/v1.0/drives/{item['parentReference']['driveId']}/items/{item['id']}/content"
    else:
        raise ValueError(f"Could not access file: {resp.status_code} {resp.text[:200]}")

    file_resp = requests.get(download_url, headers=headers, allow_redirects=True)
    if file_resp.status_code != 200:
        raise ValueError(f"Could not download file: {file_resp.status_code}")
    return file_resp.content


def get_raw_sheet(content: bytes):
    from io import BytesIO
    xl = pd.ExcelFile(BytesIO(content), engine="openpyxl")
    raw = xl.parse("Grand Summary", header=None)
    return raw, xl.sheet_names


def _fetch_sheet_via_graph(token: str) -> pd.DataFrame:
    _load_globals()
    """Use Graph Excel API to get formula-evaluated values from Grand Summary."""
    headers = {"Authorization": f"Bearer {token}"}
    sharing_url = f"u!{__import__('base64').urlsafe_b64encode(FILE_URL.encode()).decode().rstrip('=')}"

    # Get drive item ID
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{sharing_url}/driveItem",
        headers=headers,
    )
    if resp.status_code != 200:
        raise ValueError(f"Could not access file: {resp.text[:200]}")
    item = resp.json()
    drive_id = item["parentReference"]["driveId"]
    item_id  = item["id"]

    # Use Excel API to get used range with calculated values
    range_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets/Grand Summary/usedRange",
        headers=headers,
    )
    if range_resp.status_code != 200:
        raise ValueError(f"Excel API failed: {range_resp.text[:200]}")

    values = range_resp.json().get("values", [])
    return pd.DataFrame(values)


def _get_drive_item(token: str):
    _load_globals()
    """Return (headers, drive_id, item_id) for the pipeline file."""
    hdrs = {"Authorization": f"Bearer {token}"}
    sharing_url = f"u!{__import__('base64').urlsafe_b64encode(FILE_URL.encode()).decode().rstrip('=')}"
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{sharing_url}/driveItem", headers=hdrs
    )
    if resp.status_code != 200:
        raise ValueError(f"Could not access file: {resp.text[:200]}")
    item = resp.json()
    return hdrs, item["parentReference"]["driveId"], item["id"]


def fetch_client_projects(token: str, client_name: str, month: str) -> dict:
    """
    Read the client's individual sheet and return project-level data for `month`.
    Returns dict: { grand_summary_category: [{project, revenue}] }
    Maps:
      BILLED REVENUE + CONFIRMED REVENUE → "CONFIRMED"
      PROPOSED → "PROPOSED"
      ACCOUNT PLANNING / POTENTIAL → "ACCOUNT PLANNING"
      SPECULATIVE → "SPECULATIVE"
    """
    hdrs, drive_id, item_id = _get_drive_item(token)

    # URL-encode the sheet name
    sheet = requests.utils.quote(client_name, safe="")
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets/{sheet}/usedRange",
        headers=hdrs,
        timeout=6,
    )
    if resp.status_code != 200:
        return {}

    rows = resp.json().get("values", [])
    if not rows:
        return {}

    raw = pd.DataFrame(rows)

    # Find header row containing month names
    month_row_idx = None
    month_col = None
    for i, row in raw.iterrows():
        vals = [str(v).strip() if v is not None else "" for v in row]
        if month in vals:
            month_row_idx = i
            month_col = vals.index(month)
            break

    if month_row_idx is None or month_col is None:
        return {}

    # Section keyword → Grand Summary category
    SECTION_MAP = {
        "BILLED":      "BILLED",
        "CONFIRMED":   "CONFIRMED",
        "PROPOSED":    "PROPOSED",
        "POTENTIAL":   "POTENTIAL",
        "ACCOUNT":     "ACCOUNT PLANNING",
        "SPECULATIVE": "SPECULATIVE",
    }

    result = {}   # { category: {project_name: value} }
    current_cat = None

    for i in range(month_row_idx + 1, len(raw)):
        row = raw.iloc[i]
        cell = str(row.iloc[0]).strip() if row.iloc[0] is not None else ""
        cell_upper = cell.upper()

        # Detect section header
        matched_cat = next((v for k, v in SECTION_MAP.items() if k in cell_upper), None)
        if matched_cat and len(cell) > 3:
            current_cat = matched_cat
            continue

        if not current_cat:
            continue
        if not cell or cell in ("0", "£0") or "TOTAL" in cell_upper:
            continue

        # Get value for the month column
        val = _clean_value(row.iloc[month_col]) if month_col < len(row) else 0.0
        if val == 0:
            continue

        if current_cat not in result:
            result[current_cat] = {}
        # Combine BILLED + CONFIRMED for same project name
        result[current_cat][cell] = result[current_cat].get(cell, 0.0) + val

    # Convert to list of dicts
    return {
        cat: [{"Project": p, "Revenue": v} for p, v in projects.items()]
        for cat, projects in result.items()
    }


def parse_grand_summary(content: bytes, token: str = None) -> dict:
    """Parse the Grand Summary sheet into a dict of {section: DataFrame}."""
    raw = _fetch_sheet_via_graph(token) if token else None
    if raw is None:
        from io import BytesIO
        import openpyxl
        wb = openpyxl.load_workbook(BytesIO(content), data_only=True)
        ws = wb["Grand Summary"]
        data = [[cell.value for cell in row] for row in ws.iter_rows()]
        raw = pd.DataFrame(data)

    # Find which row contains "Jan" to locate month columns
    month_row_idx = None
    month_cols = {}
    for i, row in raw.iterrows():
        row_vals = [str(v).strip() if pd.notna(v) else "" for v in row]
        if "Jan" in row_vals:
            month_row_idx = i
            for m in MONTHS:
                if m in row_vals:
                    month_cols[m] = row_vals.index(m)
            break

    if not month_cols:
        raise ValueError("Could not find month columns in Grand Summary sheet")

    # Find name column (contains "CONFIRMED")
    name_col = None
    for i, row in raw.iterrows():
        for col_idx, val in enumerate(row):
            if pd.notna(val) and "CONFIRMED" in str(val).upper():
                name_col = col_idx
                break
        if name_col is not None:
            break

    if name_col is None:
        raise ValueError("Could not find name column")

    # Dynamically find section headers and collect rows until next section/total
    SECTION_KEYWORDS = {
        "CONFIRMED":       ["CONFIRMED"],
        "PROPOSED":        ["PROPOSED"],
        "ACCOUNT PLANNING":["ACCOUNT PLANNING", "POTENTIAL"],
        "SPECULATIVE":     ["SPECULATIVE"],
    }
    ALL_SECTION_KEYWORDS = [kw for kws in SECTION_KEYWORDS.values() for kw in kws]

    # Build list of (section_key, start_row, end_row) by finding section headers
    # and their corresponding TOTAL rows
    # Start from month_row_idx itself since CONFIRMED REVENUE header shares that row
    section_bounds = []
    current_sec    = None
    current_start  = None
    seen_sections  = set()  # only use first occurrence of each section

    for i in range(month_row_idx, len(raw)):
        row      = raw.iloc[i]
        cell_val = row.iloc[name_col]
        cell     = str(cell_val).strip() if pd.notna(cell_val) else ""
        cell_up  = cell.upper()

        if cell == "" or cell_up in ("NAN", "NONE"):
            continue

        # Stop scanning once we've found all 4 sections
        if len(seen_sections) == len(SECTION_KEYWORDS) and current_sec is None:
            break

        # Detect section header
        is_header = (
            any(kw in cell_up for kw in [kw for kws in SECTION_KEYWORDS.values() for kw in kws])
            and "TOTAL" not in cell_up
            and "FEE" not in cell_up
            and not any(c.isdigit() for c in cell)
            and len(cell) < 50
        )
        if is_header:
            matched = next((sec for sec, kws in SECTION_KEYWORDS.items()
                           if any(kw in cell_up for kw in kws)), None)
            if matched and matched not in seen_sections:
                if current_sec and current_start:
                    section_bounds.append((current_sec, current_start, i))
                    seen_sections.add(current_sec)
                current_sec   = matched
                current_start = i + 1
                continue
            elif matched:
                continue  # skip duplicate section headers

        # Detect total row — marks end of current section
        if "TOTAL" in cell_up and current_sec and current_start:
            section_bounds.append((current_sec, current_start, i))
            seen_sections.add(current_sec)
            current_sec   = None
            current_start = None

    if current_sec and current_start:
        section_bounds.append((current_sec, current_start, len(raw)))

    # Read rows for each section bound
    section_rows = {k: [] for k in SECTION_KEYWORDS}
    for sec_key, start, end in section_bounds:
        existing = [r["Client"] for r in section_rows[sec_key]]
        for i in range(start, end):
            if i >= len(raw):
                break
            row      = raw.iloc[i]
            cell_val = row.iloc[name_col]
            cell     = str(cell_val).strip() if pd.notna(cell_val) else ""
            cell_up  = cell.upper()
            if cell == "" or cell_up in ("NAN", "NONE"):
                continue
            if "TOTAL" in cell_up:
                break
            # Skip placeholder new clients with no revenue
            if cell_up.startswith("NEW CLIENT"):
                values_check = {m: _clean_value(row.iloc[month_cols[m]]) for m in MONTHS}
                if sum(values_check.values()) == 0:
                    continue
            values = {m: _clean_value(row.iloc[month_cols[m]]) for m in MONTHS}
            if sum(values.values()) == 0:
                continue
            # If client already exists in this section, add values
            existing_row = next((r for r in section_rows[sec_key] if r["Client"] == cell), None)
            if existing_row:
                for m in MONTHS:
                    existing_row[m] += values[m]
            else:
                section_rows[sec_key].append({"Client": cell, **values})

    result = {}
    for sec_key, rows in section_rows.items():
        if rows:
            result[sec_key] = pd.DataFrame(rows)

    return result


def fetch_pipeline(token: str) -> dict:
    _load_globals()
    content = _download_raw(token)
    return parse_grand_summary(content, token=token)


def fetch_capacity(token: str) -> pd.DataFrame:
    _load_globals()
    """
    Read Capacity full team sheet, group by department (col B), sum by month.
    Returns long-format DataFrame: Department, Month, Value
    """
    hdrs = {"Authorization": f"Bearer {token}"}
    sharing_url = f"u!{__import__('base64').urlsafe_b64encode(BUDGET_URL.encode()).decode().rstrip('=')}"

    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{sharing_url}/driveItem",
        headers=hdrs, timeout=10,
    )
    if resp.status_code != 200:
        raise ValueError(f"Cannot access budget file: {resp.text[:200]}")
    item     = resp.json()
    drive_id = item["parentReference"]["driveId"]
    item_id  = item["id"]

    sheets_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets",
        headers=hdrs, timeout=10,
    )
    sheets = sheets_resp.json().get("value", [])
    cap_sheet = next((s["name"] for s in sheets if "capacity" in s["name"].lower()), None)
    if not cap_sheet:
        raise ValueError(f"Cannot find Capacity sheet. Available: {[s['name'] for s in sheets]}")

    sheet_enc = requests.utils.quote(cap_sheet, safe="")
    range_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets/{sheet_enc}/usedRange",
        headers=hdrs, timeout=15,
    )
    if range_resp.status_code != 200:
        raise ValueError(f"Cannot read capacity sheet: {range_resp.text[:200]}")

    rows = range_resp.json().get("values", [])
    raw  = pd.DataFrame(rows)

    # Find header row with month names
    MONTH_VARIANTS = {
        "jan": "Jan", "feb": "Feb", "mar": "Mar", "apr": "Apr",
        "may": "May", "jun": "Jun", "jul": "Jul", "aug": "Aug",
        "sep": "Sep", "oct": "Oct", "nov": "Nov", "dec": "Dec",
    }
    month_row_idx = None
    month_cols    = {}
    for i, row in raw.iterrows():
        vals = [str(v).strip().lower() if v is not None else "" for v in row]
        found = {MONTH_VARIANTS[k]: idx for idx, v in enumerate(vals)
                 for k in MONTH_VARIANTS if v.startswith(k)}
        if len(found) >= 6:
            month_row_idx = i
            month_cols    = found
            break

    if not month_cols:
        raise ValueError("Cannot find month columns in Capacity sheet")

    # Find department column (col B = index 1)
    dept_col  = 1
    title_col = 0  # col A = job title
    name_col  = 4  # col E = employee name

    # Column AM (0-indexed = 38) contains staff type indicator
    AM_COL = 38

    # Manual department overrides for specific employees
    DEPT_OVERRIDES = {
        "jon forsyth":       "Creative",
        "simon massey":      "Strategy",
        "claudia wallace":   "Account Management",
    }
    # Debug: sample what's in the last few columns of the first data rows

    # Collect rows: skip headers, totals, empty departments
    SKIP_KEYWORDS = ["total", "subtotal", "grand"]
    result_rows = []

    for i in range(month_row_idx + 1, len(raw)):
        row   = raw.iloc[i]
        dept  = str(row.iloc[dept_col]).strip() if row.iloc[dept_col] is not None else ""
        title = str(row.iloc[title_col]).strip() if row.iloc[title_col] is not None else ""
        name  = str(row.iloc[name_col]).strip() if name_col < len(row) and row.iloc[name_col] is not None else ""

        if not dept or dept.lower() in ("none", "nan"):
            continue
        if any(kw in dept.lower() for kw in SKIP_KEYWORDS):
            continue
        if any(kw in name.lower() for kw in SKIP_KEYWORDS):
            continue

        # Read column AM for staff type
        am_val = str(row.iloc[AM_COL]).strip().upper() if AM_COL < len(row) and row.iloc[AM_COL] is not None else ""
        staff_type = "Freelancer" if am_val == "FREELANCE" else "Full Time"

        # Apply department overrides
        if name.lower() in DEPT_OVERRIDES:
            dept = DEPT_OVERRIDES[name.lower()]

        for m, col in month_cols.items():
            v = _clean_value(row.iloc[col] if col < len(row) else None)
            if v != 0:
                result_rows.append({"Department": dept, "Job Title": title, "Employee": name, "Staff Type": staff_type, "Month": m, "Cost": v})

    return pd.DataFrame(result_rows)


SCOPE_DEPT_MAP = {
    "MANAGEMENT":      "Management",
    "CLIENTSERVICES":  "Account Management",
    "STRATEGY":        "Strategy",
    "CREATIVE":        "Creative",
    "DESIGN":          "Design",
    "PRODUCTION":      "Production",
    "BUSINESS AFFAIRS":"Business Affairs",
}

def fetch_scope(token: str) -> pd.DataFrame:
    _load_globals()
    """
    Read each department sheet from the Scope Tracker.
    Returns long-format DataFrame: Department, Month, Chargeout
    """
    hdrs = {"Authorization": f"Bearer {token}"}
    sharing_url = f"u!{__import__('base64').urlsafe_b64encode(SCOPE_URL.encode()).decode().rstrip('=')}"

    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{sharing_url}/driveItem",
        headers=hdrs, timeout=10,
    )
    if resp.status_code != 200:
        raise ValueError(f"Cannot access scope file: {resp.text[:200]}")
    item     = resp.json()
    drive_id = item["parentReference"]["driveId"]
    item_id  = item["id"]

    sheets_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets",
        headers=hdrs, timeout=10,
    )
    sheets = sheets_resp.json().get("value", [])

    result_rows = []

    for sheet in sheets:
        dept_key = sheet["name"].upper().strip()
        if dept_key not in SCOPE_DEPT_MAP:
            continue
        dept_name   = SCOPE_DEPT_MAP[dept_key]
        sheet_enc   = requests.utils.quote(sheet["name"], safe="")

        raw_resp = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
            f"/workbook/worksheets/{sheet_enc}/usedRange",
            headers=hdrs, timeout=15,
        )
        if raw_resp.status_code != 200:
            continue

        rows = raw_resp.json().get("values", [])
        if not rows:
            continue

        # Find header row with JAN
        header_row = None
        month_cols = {}
        for i, row in enumerate(rows):
            vals = [str(v).strip().upper() if v is not None else "" for v in row]
            if "JAN" in vals:
                header_row = i
                for idx, m in enumerate(MONTHS):
                    m_upper = m.upper()
                    if m_upper in vals:
                        month_cols[m] = vals.index(m_upper)
                break

        if not month_cols or header_row is None:
            continue

        # Detect charge categories by section headers
        # Order matters: check longer/more specific strings first
        SECTION_HEADERS = [
            ("CLIENT (NOT RECOVERED)", "Client (not recovered)"),
            ("NOT RECOVERED",          "Client (not recovered)"),
            ("NEW BIZ",                "New Biz"),
            ("NEW BUSINESS",           "New Biz"),
        ]
        SKIP_LABELS = {"", "NONE", "TOTAL", "GRAND TOTAL", "PROJECT", "CLIENT"}

        current_category = "Client"  # default

        for row in rows[header_row + 1:]:
            cell = str(row[0]).strip() if row[0] is not None else ""
            cell_upper = cell.upper()

            # Always detect section headers first (col 0 keyword match)
            matched_cat = next((v for k, v in SECTION_HEADERS if k in cell_upper), None)
            if matched_cat:
                current_category = matched_cat
                continue

            # Skip column header repeat rows (PROJECT or JAN in other cols)
            row_vals_upper = [str(v).upper() if v is not None else "" for v in row]
            is_col_header = "PROJECT" in row_vals_upper[1:] or ("JAN" in row_vals_upper[1:] and cell_upper not in ("CLIENT", "CLIENT (BILLABLE)"))
            if is_col_header:
                continue

            # "CLIENT" in col 0 alone resets to Client (only if not a data row)
            if cell_upper in ("CLIENT", "CLIENT (BILLABLE)"):
                # Only reset if it looks like a section header (col 1 has "PROJECT")
                if len(row) > 1 and str(row[1]).strip().upper() == "PROJECT":
                    continue  # column header row, skip
                current_category = "Client"
                continue

            if not cell or cell_upper in SKIP_LABELS:
                continue
            if "total" in cell_upper:
                continue

            for m, col in month_cols.items():
                v = _clean_value(row[col] if col < len(row) else None)
                if v != 0:
                    result_rows.append({"Department": dept_name, "Client": cell, "Category": current_category, "Month": m, "Chargeout": v})

    return pd.DataFrame(result_rows)


def fetch_scope_structure_debug(token: str) -> list:
    """Return all rows from the MANAGEMENT sheet to see section structure."""
    hdrs = {"Authorization": f"Bearer {token}"}
    sharing_url = f"u!{__import__('base64').urlsafe_b64encode(SCOPE_URL.encode()).decode().rstrip('=')}"
    resp = requests.get(f"https://graph.microsoft.com/v1.0/shares/{sharing_url}/driveItem", headers=hdrs, timeout=10)
    item = resp.json()
    drive_id = item["parentReference"]["driveId"]
    item_id  = item["id"]
    sheets_resp = requests.get(f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets", headers=hdrs, timeout=10)
    sheets = sheets_resp.json().get("value", [])
    # Get MANAGEMENT sheet
    mgmt = next((s for s in sheets if s["name"].upper() == "MANAGEMENT"), sheets[1])
    sheet_enc = requests.utils.quote(mgmt["name"], safe="")
    raw_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets/{sheet_enc}/usedRange",
        headers=hdrs, timeout=15,
    )
    rows = raw_resp.json().get("values", [])
    # Return first col of every row to see section headers
    return [{"row": i, "col_A": str(r[0]) if r else ""} for i, r in enumerate(rows)]


def fetch_scope_debug(token: str) -> dict:
    """Fetch raw data from first sheet of scope tracker for debugging."""
    hdrs = {"Authorization": f"Bearer {token}"}
    sharing_url = f"u!{__import__('base64').urlsafe_b64encode(SCOPE_URL.encode()).decode().rstrip('=')}"
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{sharing_url}/driveItem",
        headers=hdrs, timeout=10,
    )
    if resp.status_code != 200:
        raise ValueError(f"Cannot access scope file: {resp.text[:200]}")
    item     = resp.json()
    drive_id = item["parentReference"]["driveId"]
    item_id  = item["id"]

    sheets_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets",
        headers=hdrs, timeout=10,
    )
    sheets = sheets_resp.json().get("value", [])
    sheet_names = [s["name"] for s in sheets]

    # Fetch first sheet raw
    first_sheet = requests.utils.quote(sheets[0]["name"], safe="")
    raw_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets/{first_sheet}/usedRange",
        headers=hdrs, timeout=15,
    )
    rows = raw_resp.json().get("values", [])
    return {"sheet_names": sheet_names, "first_sheet": sheets[0]["name"], "rows": rows[:15]}


def fetch_raw(token: str):
    raw = _fetch_sheet_via_graph(token)
    return raw, ["Grand Summary"]


def fetch_budget(token: str) -> dict:
    _load_globals()
    """
    Fetch the P&L from the Budget file (2nd sheet: Summary).
    Returns dict with keys: Total income, Total staff costs,
    Total office & Other costs, Net Profit, Operating Profit Margin
    Each value is a dict of {month: value}.
    """
    hdrs = {"Authorization": f"Bearer {token}"}
    sharing_url = f"u!{__import__('base64').urlsafe_b64encode(BUDGET_URL.encode()).decode().rstrip('=')}"

    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{sharing_url}/driveItem",
        headers=hdrs, timeout=10,
    )
    if resp.status_code != 200:
        raise ValueError(f"Cannot access budget file: {resp.text[:200]}")
    item     = resp.json()
    drive_id = item["parentReference"]["driveId"]
    item_id  = item["id"]

    # Get sheet list to find 2nd sheet name
    sheets_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets",
        headers=hdrs, timeout=10,
    )
    if sheets_resp.status_code != 200:
        raise ValueError(f"Cannot list sheets: {sheets_resp.text[:200]}")
    sheets = sheets_resp.json().get("value", [])
    sheet_names = [s["name"] for s in sheets]
    # Find "Summary P&L" sheet by name
    target = next((s["name"] for s in sheets if "summary p&l" in s["name"].lower() or "summary pl" in s["name"].lower()), None)
    if not target:
        target = next((s["name"] for s in sheets if "summary" in s["name"].lower() and "p" in s["name"].lower()), None)
    if not target:
        raise ValueError(f"Cannot find Summary P&L sheet. Available: {sheet_names}")
    summary_sheet = requests.utils.quote(target, safe="")

    range_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets/{summary_sheet}/usedRange",
        headers=hdrs, timeout=15,
    )
    if range_resp.status_code != 200:
        raise ValueError(f"Cannot read summary sheet: {range_resp.text[:200]}")

    rows = range_resp.json().get("values", [])
    raw  = pd.DataFrame(rows)

    # Find header row with month names — handles "Jan", "JAN/25", "Jan-25" etc.
    MONTH_VARIANTS = {
        "jan": "Jan", "feb": "Feb", "mar": "Mar", "apr": "Apr",
        "may": "May", "jun": "Jun", "jul": "Jul", "aug": "Aug",
        "sep": "Sep", "oct": "Oct", "nov": "Nov", "dec": "Dec",
    }
    month_row_idx = None
    month_cols    = {}
    for i, row in raw.iterrows():
        vals = [str(v).strip().lower() if v is not None else "" for v in row]
        found = {MONTH_VARIANTS[k]: idx for idx, v in enumerate(vals)
                 for k in MONTH_VARIANTS if v.startswith(k)}
        if len(found) >= 6:
            month_row_idx = i
            month_cols    = found
            break

    if not month_cols:
        sample = [[str(v) for v in raw.iloc[i].tolist()[:16]] for i in range(min(12, len(raw)))]
        raise ValueError(f"Cannot find month columns. First 12 rows: {sample}")

    # Rows we care about
    TARGET_ROWS = {
        "Total income":                 "total turnover",
        "Total staff costs":            "total staff",
        "Total office & Other costs":   "total office",
        "Net Profit":                   "net profit",
        "Operating Profit Margin":      "operating profit margin",
    }

    # Find which column contains the P&L summary labels
    label_col = 0
    for i in range(month_row_idx + 1, min(month_row_idx + 60, len(raw))):
        for col_idx in range(min(month_cols.get("Jan", 2), len(raw.columns))):
            val = raw.iloc[i, col_idx]
            if val and "total income" in str(val).lower():
                label_col = col_idx
                break
        if label_col != 0:
            break

    result = {k: {} for k in TARGET_ROWS}

    for i in range(month_row_idx + 1, len(raw)):
        row      = raw.iloc[i]
        cell_val = row.iloc[label_col] if label_col < len(row) else ""
        cell     = str(cell_val).strip().lower() if cell_val is not None else ""

        matched = next((k for k, kw in TARGET_ROWS.items() if kw in cell), None)
        if not matched or matched == "Total office & Other costs":
            continue

        for m, col in month_cols.items():
            val = row.iloc[col] if col < len(row) else None
            if matched == "Operating Profit Margin":
                try:
                    v = float(str(val).replace("%", "").strip()) if val not in (None, "", " ") else 0.0
                    result[matched][m] = v if abs(v) <= 1 else v / 100
                except (ValueError, TypeError):
                    result[matched][m] = 0.0
            else:
                result[matched][m] = _clean_value(val)

    # ── Office & Other costs from Detail P&L 24, rows 35-68, excl. Staff Costs ──
    detail_sheet_name = next((s["name"] for s in sheets if "detail p&l" in s["name"].lower()), None)
    if detail_sheet_name:
        detail_sheet = requests.utils.quote(detail_sheet_name, safe="")
        det_resp = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
            f"/workbook/worksheets/{detail_sheet}/usedRange",
            headers=hdrs, timeout=15,
        )
        if det_resp.status_code == 200:
            det_rows = det_resp.json().get("values", [])
            det_raw  = pd.DataFrame(det_rows)

            # Find month columns in this sheet
            det_month_cols = {}
            det_month_row  = None
            for i, row in det_raw.iterrows():
                vals = [str(v).strip().lower() if v is not None else "" for v in row]
                found = {MONTH_VARIANTS[k]: idx for idx, v in enumerate(vals)
                         for k in MONTH_VARIANTS if v.startswith(k)}
                if len(found) >= 6:
                    det_month_row  = i
                    det_month_cols = found
                    break

            if det_month_cols:
                # Find column Z index (Category 2) — look for "category" header
                cat_col = None
                header_row = det_raw.iloc[det_month_row] if det_month_row is not None else None
                if header_row is not None:
                    for ci, v in enumerate(header_row):
                        if v and "category" in str(v).lower():
                            cat_col = ci

                # Sum rows 35-68 (0-indexed: 34-67) excluding Staff Costs
                office_totals = {m: 0.0 for m in MONTHS}
                jan_debug = []
                cos_jan = 0.0

                # Read rows 24 and 54 directly by Excel address (col I = Jan, J=Feb...T=Dec)
                EXCEL_MONTH_COLS = "IJKLMNOPQRST"  # I=Jan through T=Dec
                for deduct_row in [24, 54]:
                    addr = f"I{deduct_row}:T{deduct_row}"
                    cell_resp = requests.get(
                        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
                        f"/workbook/worksheets/{detail_sheet}/range(address='{addr}')",
                        headers=hdrs, timeout=10,
                    )
                    if cell_resp.status_code == 200:
                        vals = cell_resp.json().get("values", [[]])[0]
                        for idx, m in enumerate(MONTHS):
                            v = _clean_value(vals[idx] if idx < len(vals) else None)
                            office_totals[m] -= v
                            if m == "Jan":
                                cos_jan += v
                                jan_debug.append({"row": deduct_row, "name": f"DEDUCTED row {deduct_row}", "cat": "—", "Jan": -v})

                for i in list(range(34, min(68, len(det_raw)))) + ([73] if len(det_raw) > 73 else []):
                    row = det_raw.iloc[i]
                    cat  = str(row.iloc[cat_col]).strip() if cat_col is not None and cat_col < len(row) else ""
                    name = str(row.iloc[0]).strip() if row.iloc[0] is not None else ""
                    if not cat or cat.lower() in ("none", "nan") or "staff" in cat.lower():
                        continue
                    for m, col in det_month_cols.items():
                        v = _clean_value(row.iloc[col] if col < len(row) else None)
                        office_totals[m] += v
                        if m == "Jan" and v != 0:
                            jan_debug.append({"row": i+1, "name": name, "cat": cat, "Jan": v})

                result["Total office & Other costs"] = office_totals
                result["__office_debug__"] = jan_debug + [{"row": 24, "name": "DEDUCTED: Total cost of sales", "cat": "—", "Jan": -cos_jan}]
                result["__office_total_jan__"] = office_totals.get("Jan", 0)

    return result
