"""
tickers.py

Owns the tracked company universe. Every other module (financials.py,
valuation.py, news.py, analysis.py, excel_workbook.py, main.py) calls into
this module to find out which tickers exist instead of hardcoding a list.

The universe lives in data/tickers.json, not in Python source. That is
what makes "add a ticker -> zero code changes" possible: add_company.py
calls add_ticker() at runtime, which appends a record to that JSON file.
Every module that reads get_all_tickers() picks the new company up on its
next run automatically.

Removal is soft by default (active=False): historical financials, ratios,
and news already written for a ticker must never disappear from the sheet
just because it's no longer tracked going forward.
"""

import json
import logging
import os
from datetime import date

from config import TICKERS_FILE
from utils import WRITE_LOCK

logger = logging.getLogger(__name__)

# Bootstraps data/tickers.json on first run only. After that, the JSON
# file is the sole source of truth -- this constant is never read again.
_SEED_UNIVERSE = [
    {"ticker": "AAPL", "name": "Apple Inc.", "sector": "Technology"},
    {"ticker": "MSFT", "name": "Microsoft Corporation", "sector": "Technology"},
    {"ticker": "NVDA", "name": "NVIDIA Corporation", "sector": "Semiconductors"},
    {"ticker": "TSM", "name": "Taiwan Semiconductor Manufacturing Company", "sector": "Semiconductors"},
    {"ticker": "AVGO", "name": "Broadcom Inc.", "sector": "Semiconductors"},
    {"ticker": "JPM", "name": "JPMorgan Chase & Co.", "sector": "Financial Services"},
    {"ticker": "V", "name": "Visa Inc.", "sector": "Financial Services"},
    {"ticker": "MA", "name": "Mastercard Incorporated", "sector": "Financial Services"},
    {"ticker": "AMZN", "name": "Amazon.com, Inc.", "sector": "Consumer Discretionary"},
    {"ticker": "MCD", "name": "McDonald's Corporation", "sector": "Consumer Discretionary"},
    {"ticker": "NKE", "name": "Nike, Inc.", "sector": "Consumer Discretionary"},
    {"ticker": "COST", "name": "Costco Wholesale Corporation", "sector": "Consumer Staples"},
    {"ticker": "KO", "name": "The Coca-Cola Company", "sector": "Consumer Staples"},
    {"ticker": "PG", "name": "The Procter & Gamble Company", "sector": "Consumer Staples"},
    {"ticker": "LLY", "name": "Eli Lilly and Company", "sector": "Healthcare"},
    {"ticker": "UNH", "name": "UnitedHealth Group Incorporated", "sector": "Healthcare"},
    {"ticker": "CAT", "name": "Caterpillar Inc.", "sector": "Industrials"},
    {"ticker": "HON", "name": "Honeywell International Inc.", "sector": "Industrials"},
    {"ticker": "XOM", "name": "Exxon Mobil Corporation", "sector": "Energy"},
    {"ticker": "NEE", "name": "NextEra Energy, Inc.", "sector": "Utilities"},
    {"ticker": "GOOGL", "name": "Alphabet Inc.", "sector": "Communication Services"},
    {"ticker": "META", "name": "Meta Platforms, Inc.", "sector": "Communication Services"},
    {"ticker": "PLD", "name": "Prologis, Inc.", "sector": "Real Estate"},
    {"ticker": "SHW", "name": "The Sherwin-Williams Company", "sector": "Materials"},
]


def _seed_record(entry: dict) -> dict:
    return {
        "ticker": entry["ticker"],
        "name": entry["name"],
        "sector": entry["sector"],
        "added_date": date.today().isoformat(),
        "active": True,
    }


def _atomic_write(records: list[dict]) -> None:
    """Write JSON via temp-file + os.replace so a crash mid-write can't
    corrupt the ticker file -- this file is the app's source of truth."""
    tmp_path = f"{TICKERS_FILE}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2)
    os.replace(tmp_path, TICKERS_FILE)


def load_tickers() -> list[dict]:
    """Return every ticker record (active and inactive). Seeds the file
    with the starting universe the first time it's called."""
    if not TICKERS_FILE.exists():
        logger.info("No tickers.json found -- seeding with starting universe (%d companies)", len(_SEED_UNIVERSE))
        records = [_seed_record(e) for e in _SEED_UNIVERSE]
        _atomic_write(records)
        return records

    try:
        with open(TICKERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError) as exc:
        logger.error("Failed to read %s: %s", TICKERS_FILE, exc)
        raise


def save_tickers(records: list[dict]) -> None:
    _atomic_write(records)


def get_ticker_records(active_only: bool = True) -> list[dict]:
    records = load_tickers()
    if active_only:
        return [r for r in records if r.get("active", True)]
    return records


def get_all_tickers(active_only: bool = True) -> list[str]:
    """The list every collection module (financials.py, news.py, ...)
    should iterate over."""
    return [r["ticker"] for r in get_ticker_records(active_only=active_only)]


def get_ticker_record(ticker: str) -> dict | None:
    ticker = ticker.upper()
    for r in load_tickers():
        if r["ticker"] == ticker:
            return r
    return None


def ticker_exists(ticker: str) -> bool:
    return get_ticker_record(ticker) is not None


def get_sectors() -> list[str]:
    return sorted({r["sector"] for r in get_ticker_records(active_only=True)})


def get_tickers_by_sector(sector: str) -> list[dict]:
    return [r for r in get_ticker_records(active_only=True) if r["sector"] == sector]


def add_ticker(ticker: str, name: str, sector: str) -> dict:
    """Append a new company to the universe. Called by add_company.py
    after the ticker has been validated and its data downloaded.

    Raises ValueError if the ticker is already tracked (including
    inactive/removed ones -- reactivate() should be used for those).
    """
    ticker = ticker.upper()
    with WRITE_LOCK:
        records = load_tickers()

        existing = next((r for r in records if r["ticker"] == ticker), None)
        if existing is not None:
            raise ValueError(
                f"{ticker} is already in the universe (active={existing.get('active', True)})"
            )

        record = _seed_record({"ticker": ticker, "name": name, "sector": sector})
        records.append(record)
        save_tickers(records)
    logger.info("Added %s (%s, %s) to the universe", ticker, name, sector)
    return record


def remove_ticker(ticker: str, hard: bool = False) -> None:
    """Soft-remove by default: marks the ticker inactive so future runs
    skip it, but its historical rows in the Excel workbook are untouched.
    hard=True deletes the record entirely (does not touch already-written
    workbook data, which lives in the .xlsx file, not this file)."""
    ticker = ticker.upper()
    with WRITE_LOCK:
        records = load_tickers()

        if hard:
            new_records = [r for r in records if r["ticker"] != ticker]
            if len(new_records) == len(records):
                raise ValueError(f"{ticker} is not in the universe")
            save_tickers(new_records)
            logger.info("Hard-removed %s from the universe", ticker)
            return

        record = next((r for r in records if r["ticker"] == ticker), None)
        if record is None:
            raise ValueError(f"{ticker} is not in the universe")
        record["active"] = False
        save_tickers(records)
        logger.info("Deactivated %s (soft remove)", ticker)


def reactivate_ticker(ticker: str) -> None:
    ticker = ticker.upper()
    with WRITE_LOCK:
        records = load_tickers()
        record = next((r for r in records if r["ticker"] == ticker), None)
        if record is None:
            raise ValueError(f"{ticker} is not in the universe")
        record["active"] = True
        save_tickers(records)
    logger.info("Reactivated %s", ticker)
