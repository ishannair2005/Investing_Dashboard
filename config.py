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
