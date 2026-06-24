from __future__ import annotations

import argparse
import json
import os
from pathlib import Path

from .config import settings
from .service import commit_outputs, update_dashboard


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate MVIKAS dashboard script.js from Excel/Google Sheet data.")
    group = parser.add_mutually_exclusive_group(required=False)
    group.add_argument("--source-url", help="Google Sheet / XLSX URL to download")
    group.add_argument("--excel", help="Local .xlsx file path")
    parser.add_argument("--static-dir", default=os.getenv("MVIKAS_STATIC_DIR", str(settings.static_dir)), help="Directory where script.js will be written")
    parser.add_argument("--report-date", default=None, help="Optional override date, e.g. 2026-06-18")
    parser.add_argument("--commit", action="store_true", help="Commit and push generated files using git settings")
    args = parser.parse_args()

    source = args.excel or args.source_url or settings.source_url
    if not source:
        raise SystemExit("Provide --excel, --source-url, or MVIKAS_SOURCE_URL in .env")

    result = update_dashboard(source=source, static_dir=Path(args.static_dir), report_date=args.report_date)
    if args.commit:
        commit_outputs(settings)
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
