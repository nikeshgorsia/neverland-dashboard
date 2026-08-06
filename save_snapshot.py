"""
Scheduled snapshot script — run by GitHub Actions every Thursday at 6pm GMT.
Fetches live pipeline from SharePoint using app-only auth and saves a dated snapshot.
"""
import sys
from sharepoint_sync import get_app_token, fetch_pipeline, save_pipeline_snapshot

def main():
    print("Acquiring app token...")
    token = get_app_token()
    print("Fetching pipeline...")
    pipeline = fetch_pipeline(token)
    if not pipeline:
        print("ERROR: fetch_pipeline returned empty data", file=sys.stderr)
        sys.exit(1)
    date_str = save_pipeline_snapshot(pipeline)
    print(f"Snapshot saved: {date_str}")

if __name__ == "__main__":
    main()
