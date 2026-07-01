"""
Runs via GitHub Actions every Thursday at 6pm GMT.
Uses app-only (client credentials) auth — no browser needed.
Requires Azure AD app to have Files.Read.All application permission.
"""
import os, sys, json, base64, datetime, requests
import pandas as pd
from msal import ConfidentialClientApplication

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

TENANT_ID    = os.environ["TENANT_ID"]
CLIENT_ID    = os.environ["CLIENT_ID"]
CLIENT_SECRET= os.environ["CLIENT_SECRET"]
FILE_URL     = os.environ["SHAREPOINT_FILE_URL"]
GITHUB_TOKEN = os.environ["GITHUB_TOKEN"]
GITHUB_REPO  = "nikeshgorsia/neverland-dashboard"

MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
SECTIONS = ["CONFIRMED","PROPOSED","ACCOUNT PLANNING","SPECULATIVE"]


def get_token():
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise ValueError(f"Auth failed: {result.get('error_description', result)}")
    return result["access_token"]


def fetch_pipeline(token):
    from sharepoint_sync import parse_grand_summary, _download_raw
    # Temporarily patch FILE_URL into environment
    import sharepoint_sync as sp
    sp.FILE_URL = FILE_URL
    content = _download_raw(token)
    return parse_grand_summary(content, token=token)


def save_snapshot(pipeline_data, date_str):
    serializable = {}
    for k, v in pipeline_data.items():
        if isinstance(v, pd.DataFrame):
            serializable[k] = v.to_dict(orient="records")

    content_str = json.dumps({"date": date_str, "data": serializable}, indent=2)
    encoded = base64.b64encode(content_str.encode()).decode()
    filename = f"snapshots/snapshot_{date_str}.json"

    hdrs = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
    }
    check = requests.get(
        f"https://api.github.com/repos/{GITHUB_REPO}/contents/{filename}",
        headers=hdrs, timeout=10,
    )
    sha = check.json().get("sha") if check.status_code == 200 else None

    payload = {"message": f"Weekly snapshot {date_str}", "content": encoded, "branch": "main"}
    if sha:
        payload["sha"] = sha

    resp = requests.put(
        f"https://api.github.com/repos/{GITHUB_REPO}/contents/{filename}",
        headers=hdrs, json=payload, timeout=15,
    )
    resp.raise_for_status()
    print(f"Snapshot saved: {filename}")


if __name__ == "__main__":
    print("Fetching token...")
    token = get_token()
    print("Fetching pipeline...")
    pipeline = fetch_pipeline(token)
    today = datetime.date.today()
    monday = today - datetime.timedelta(days=today.weekday())
    date_str = monday.isoformat()
    print(f"Saving snapshot {date_str}...")
    save_snapshot(pipeline, date_str)
    print("Done.")
