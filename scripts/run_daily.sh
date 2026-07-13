#!/bin/bash
# Wrapper launchd calls for the daily automation (news + AI thesis
# refresh). Resolves the project root relative to this script's own
# location, not the caller's cwd, since launchd invokes scripts with an
# unpredictable working directory.
set -euo pipefail
PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$PROJECT_DIR"
source venv/bin/activate
# caffeinate -i: keep the system from idle-sleeping for the duration of
# this command specifically -- a run that takes 20-40+ minutes across
# 25+ AI-backed tickers can otherwise get suspended mid-request by the
# Mac idling to sleep, which showed up as multi-hour gaps in the logs
# with no way for our own request-timeout handling to help (the whole
# process is paused at the OS level, not just one network call). Can't
# override a manually closed lid, only idle sleep -- the much more
# likely case for an unattended scheduled run.
caffeinate -i python3 main.py --daily
