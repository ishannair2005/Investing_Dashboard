"""
utils.py

Shared helpers used across every data-collection module:
  - retry(): exponential-backoff retry for flaky network calls
    (yfinance and AI provider APIs both fail intermittently).
  - validate_ticker(): confirms a symbol is real before add_company.py
    downloads anything for it.
  - safe_divide() / safe_get(): financial data is full of missing or
    zero fields (zero debt, no dividend, delisted comps, etc.) --
    ratios.py and valuation.py lean on these instead of each writing
    their own None/ZeroDivisionError guards.
  - quarter_label(): the canonical "Q<n> YYYY" key financials.py uses
    to detect whether a quarter is already stored, so it never
    overwrites history.
"""

import functools
import json
import logging
import re
import time
from datetime import datetime
from typing import Any, Callable, Optional, TypeVar

import pandas as pd
import yfinance as yf

from config import MAX_RETRIES, RETRY_BACKOFF_SECONDS

logger = logging.getLogger(__name__)

T = TypeVar("T")


def retry(
    max_retries: int = MAX_RETRIES,
    backoff_seconds: float = RETRY_BACKOFF_SECONDS,
    exceptions: tuple = (Exception,),
) -> Callable:
    """Retry a function with exponential backoff, then re-raise.

    Every module that calls out to yfinance or an AI provider API wraps
    the call site with this instead of writing its own retry loop, so
    backoff behavior is tuned in exactly one place (config.py).
    """

    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        @functools.wraps(func)
        def wrapper(*args, **kwargs) -> T:
            last_exc: Optional[Exception] = None
            for attempt in range(1, max_retries + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as exc:
                    last_exc = exc
                    if attempt == max_retries:
                        break
                    wait = backoff_seconds * (2 ** (attempt - 1))
                    logger.warning(
                        "%s failed (attempt %d/%d): %s -- retrying in %.1fs",
                        func.__name__, attempt, max_retries, exc, wait,
                    )
                    time.sleep(wait)
            logger.error("%s failed after %d attempts: %s", func.__name__, max_retries, last_exc)
            raise last_exc

        return wrapper

    return decorator


@retry(exceptions=(Exception,))
def get_yf_ticker(ticker: str) -> yf.Ticker:
    """Fetch a yfinance Ticker object, retried on transient failure."""
    return yf.Ticker(ticker)


@retry(exceptions=(Exception,))
def _fetch_ticker_info(ticker: str) -> dict:
    # yf.Ticker() construction is lazy -- the actual HTTP request (and
    # thus the actual risk of a transient failure) happens on .info
    # access, so the retry has to wrap *this*, not just the constructor.
    return yf.Ticker(ticker).info or {}


def validate_ticker(ticker: str) -> bool:
    """Confirm a symbol resolves to a real, currently priced company.

    yfinance does not raise for unknown symbols -- it returns a
    near-empty info dict. So "valid" here means: has an identifying
    name AND has a current price, not just "didn't throw".

    Retries transient failures the same as every other yfinance call in
    this codebase. Without this, a single dropped request (a cold
    container's first-ever request to Yahoo Finance, a momentary rate
    limit) looks identical to "this ticker doesn't exist", which is
    actively misleading for a ticker that's actually real.
    """
    ticker = ticker.strip().upper()
    if not ticker:
        return False
    try:
        info = _fetch_ticker_info(ticker)
    except Exception as exc:
        logger.warning(
            "Ticker validation for %s failed after %d attempts -- this may be a temporary "
            "network or rate-limit issue with the data provider rather than an invalid ticker: %s",
            ticker, MAX_RETRIES, exc,
        )
        return False

    has_identity = bool(info.get("longName") or info.get("shortName"))
    has_price = info.get("currentPrice") is not None or info.get("regularMarketPrice") is not None
    return has_identity and has_price


def safe_divide(numerator: Optional[float], denominator: Optional[float]) -> Optional[float]:
    """Divide, returning None instead of raising on missing/zero inputs.

    Treats pandas/NumPy NaN the same as None: a DataFrame cell for a
    metric a company doesn't report comes back as NaN, not None, and
    without this check it would silently propagate into every ratio
    derived from it instead of yielding a clean "unavailable".
    """
    if numerator is None or denominator is None:
        return None
    if pd.isna(numerator) or pd.isna(denominator):
        return None
    try:
        if denominator == 0:
            return None
        return numerator / denominator
    except (TypeError, ZeroDivisionError):
        return None


def safe_get(d: Optional[dict], *keys: str, default: Any = None) -> Any:
    """Return the first present, non-None value for any of `keys`.

    yfinance's .info dict uses inconsistent field names across tickers
    and library versions (e.g. 'currentPrice' vs 'regularMarketPrice'),
    so callers pass fallbacks in priority order.
    """
    if not d:
        return default
    for key in keys:
        value = d.get(key)
        if value is not None:
            return value
    return default


def pct_change(current: Optional[float], previous: Optional[float]) -> Optional[float]:
    """Percentage change from previous -> current, as a decimal (0.10 = 10%)."""
    if current is None or previous is None:
        return None
    if pd.isna(current) or pd.isna(previous) or previous == 0:
        return None
    return (current - previous) / abs(previous)


def quarter_label(period_end: Any) -> str:
    """Format a period-end date as 'Q<n> YYYY'.

    This is the key financials.py checks against already-stored rows
    to decide whether a quarter is new -- it's what makes "append new
    quarters, never overwrite history" possible.
    """
    if not isinstance(period_end, (pd.Timestamp, datetime)):
        period_end = pd.to_datetime(period_end)
    quarter = (period_end.month - 1) // 3 + 1
    return f"Q{quarter} {period_end.year}"


def today_str() -> str:
    return datetime.today().strftime("%Y-%m-%d")


def extract_json_object(text: str) -> dict:
    """Pull the first {...} blob out of an AI response and parse it.

    Shared by news.py (per-headline classification) and analysis.py
    (full investment writeups) -- both ask the AI provider for a JSON
    object and both need to tolerate markdown code fences or stray
    commentary the model wraps around it instead of raw JSON.
    """
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        raise ValueError(f"No JSON object found in AI response: {text[:200]!r}")
    return json.loads(match.group(0))
