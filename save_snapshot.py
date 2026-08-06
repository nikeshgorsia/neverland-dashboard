"""
Scheduled snapshot script — run by GitHub Actions every Thursday at 6pm GMT.
Uses the cached MSAL token (TOKEN_CACHE secret) to authenticate with SharePoint.
"""
import os, sys

# Write TOKEN_CACHE env var to a temp file so _load_cache() can pick it up
_token_cache_str = os.getenv("TOKEN_CACHE", "")
if _token_cache_str:
    import tempfile
    _cache_path = os.path.expanduser("~/.neverland_token_cache.bin")
    with open(_cache_path, "w") as _f:
        _f.write(_token_cache_str)

from sharepoint_sync import get_device_flow, fetch_pipeline, save_pipeline_snapshot

def main():
    print("Acquiring token from cache...")
    app, flow, token = get_device_flow()
    if not token:
        print("ERROR: No valid cached token. Re-authenticate via the dashboard first.", file=sys.stderr)
        sys.exit(1)
    print("Fetching pipeline...")
    pipeline = fetch_pipeline(token)
    if not pipeline:
        print("ERROR: fetch_pipeline returned empty data", file=sys.stderr)
        sys.exit(1)
    date_str = save_pipeline_snapshot(pipeline)
    print(f"Snapshot saved: {date_str}")

if __name__ == "__main__":
    main()
