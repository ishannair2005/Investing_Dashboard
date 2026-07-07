#!/bin/bash
# Wrapper launchd calls for the daily automation (news + AI thesis
# refresh). Resolves the project root relative to this script's own
# location, not the caller's cwd, since launchd invokes scripts with an
# unpredictable working directory.
set -euo pipefail
PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$PROJECT_DIR"
source venv/bin/activate
python3 main.py --daily
