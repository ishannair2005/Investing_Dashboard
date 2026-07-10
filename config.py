"""
config.py

Central configuration for the Investment Research Dashboard.

Every other module imports settings from here instead of reading
os.environ or hardcoding paths directly. This keeps credentials,
file locations, and tunable behavior (retries, timeouts, logging)
in exactly one place.
"""

import logging
import logging.handlers
import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

# --------------------------------------------------------------------------
# Paths
# --------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
LOG_DIR = BASE_DIR / "logs"

DATA_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

# Single source of truth for the tracked company universe (owned by tickers.py)
TICKERS_FILE = DATA_DIR / "tickers.json"

# --------------------------------------------------------------------------
# Output workbook
# --------------------------------------------------------------------------
# A local .xlsx file, not Google Sheets -- no cloud credentials, sharing,
# or API quotas needed. Defaults to the project root; override via .env
# to point at, say, a synced Dropbox/iCloud folder.
#
# Uses `or` rather than os.getenv's default= param: .env.example ships
# with "EXCEL_FILE_PATH=" (present but empty) as a documented "unset"
# placeholder, and os.getenv only falls back to default when the key is
# absent entirely -- an empty string is still a value, so default=
# alone would resolve this to Path(""), i.e. the current directory.
EXCEL_FILE_PATH = Path(os.getenv("EXCEL_FILE_PATH") or (BASE_DIR / "Investment_Research_Dashboard.xlsx"))

# --------------------------------------------------------------------------
# Deployment mode
# --------------------------------------------------------------------------
# True on your local Mac, where the sync automation runs and the workbook
# is the live, writable source of truth. The Streamlit Community Cloud
# deployment sets this to "false" via its Secrets panel: that instance's
# filesystem is ephemeral (wiped on every restart) and has no cached git
# credentials, so on its own it couldn't persist a write -- see
# GITHUB_TOKEN below for how the cloud instance can still write, safely.
IS_LOCAL_INSTANCE = os.getenv("IS_LOCAL_INSTANCE", "true").lower() == "true"

# A GitHub fine-grained Personal Access Token (repo-scoped, Contents:
# read/write only), set as a Secret on the Streamlit Community Cloud
# deployment specifically so that instance can commit+push on its own --
# it has no local git credential cache the way your Mac does. Locally,
# leave this unset; git_sync.py falls back to a plain `git push` using
# your already-authenticated local credentials.
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
GITHUB_REPO = os.getenv("GITHUB_REPO", "ishannair2005/Investing_Dashboard")

# Whether dashboard_app.py shows write actions (add/remove company, edit
# Watchlist) at all: yes locally, and yes on any instance -- including a
# deployed one -- that has a way to persist the change back to GitHub.
CAN_EDIT_REMOTELY = IS_LOCAL_INSTANCE or bool(GITHUB_TOKEN)

# --------------------------------------------------------------------------
# AI provider (used by analysis.py)
#
# Provider-agnostic on purpose: analysis.py will select an implementation
# based on AI_PROVIDER, so switching providers later is a config change,
# not a code change. Both key slots are optional here since only one
# provider's key needs to be set at a time.
# --------------------------------------------------------------------------
AI_PROVIDER = os.getenv("AI_PROVIDER", "anthropic")  # "anthropic" | "openai"
AI_MODEL = os.getenv("AI_MODEL", "claude-opus-4-8")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

# Separate from REQUEST_TIMEOUT_SECONDS below: both provider SDKs default
# to a multi-minute internal timeout per attempt, and the full investment
# thesis (max_tokens=4096) can legitimately take a while to generate, so
# this needs more headroom than the lightweight yfinance-style calls
# REQUEST_TIMEOUT_SECONDS is tuned for. Still far short of the SDKs' own
# defaults -- bounds a stalled request (e.g. a laptop waking from sleep
# mid-call) to a few minutes total across retries instead of 20-30+.
AI_REQUEST_TIMEOUT_SECONDS = int(os.getenv("AI_REQUEST_TIMEOUT_SECONDS", 120))

# --------------------------------------------------------------------------
# Network / retry behavior (shared by financials.py, valuation.py, news.py, ...)
# --------------------------------------------------------------------------
MAX_RETRIES = int(os.getenv("MAX_RETRIES", 3))
RETRY_BACKOFF_SECONDS = float(os.getenv("RETRY_BACKOFF_SECONDS", 2.0))
REQUEST_TIMEOUT_SECONDS = int(os.getenv("REQUEST_TIMEOUT_SECONDS", 30))

# --------------------------------------------------------------------------
# News
# --------------------------------------------------------------------------
NEWS_LOOKBACK_DAYS = int(os.getenv("NEWS_LOOKBACK_DAYS", 7))
MAX_NEWS_ITEMS_PER_TICKER = int(os.getenv("MAX_NEWS_ITEMS_PER_TICKER", 10))

# --------------------------------------------------------------------------
# Logging
# --------------------------------------------------------------------------
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()


def configure_logging() -> None:
    """Configure the root logger once: console output + rotating log file.

    Guarded by root_logger.handlers so re-importing config (tests, notebooks,
    repeated module imports) never double-attaches handlers and duplicates
    every log line.
    """
    root_logger = logging.getLogger()
    if root_logger.handlers:
        return

    root_logger.setLevel(LOG_LEVEL)

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    file_handler = logging.handlers.RotatingFileHandler(
        LOG_DIR / "dashboard.log",
        maxBytes=5 * 1024 * 1024,  # 5 MB
        backupCount=5,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)


configure_logging()
