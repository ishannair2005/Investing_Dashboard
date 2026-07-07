#!/bin/bash
# Wrapper launchd calls for the "quarterly" automation (financials +
# ratios + everything else). Scheduled weekly, not literally quarterly --
# sync_financials()/sync_ratios() are cheap no-ops when no new quarter
# has been reported, so running this weekly just means a new earnings
# release gets picked up within days instead of waiting up to 3 months
# for a true quarterly cron tick.
set -euo pipefail
PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$PROJECT_DIR"
source venv/bin/activate
python3 main.py --quarterly
