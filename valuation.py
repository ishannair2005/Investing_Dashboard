"""
valuation.py

Pulls point-in-time valuation metrics and company profile information
from Yahoo Finance (yfinance's .info payload). Both come from the same
underlying API call, so this module fetches .info once per ticker and
serves two views of it:

  - get_valuation_snapshot(ticker): current price, market cap, EV,
    multiples, dividend yield, beta, 52-week range, volume.
  - get_company_profile(ticker): name, sector, industry, description,
    and other descriptive fields for the Dashboard/Watchlist tabs.

Unlike financials.py, a valuation snapshot is a current point-in-time
reading, not a permanent historical record -- callers should overwrite
each ticker's row in the Valuation tab on every run rather than append.
Each snapshot is timestamped (as_of) in case a dated log is wanted later.
"""

import logging
from datetime import datetime

from utils import get_yf_ticker, retry, safe_divide, safe_get

logger = logging.getLogger(__name__)

# Ordered list of every valuation metric this module returns (excludes
# ticker/as_of) -- the single source of truth for the Valuation tab's
# column headers in excel_workbook.py.
VALUATION_METRICS = [
    "price", "market_cap", "enterprise_value", "shares_outstanding",
    "pe_ratio", "forward_pe", "peg_ratio", "price_to_sales", "price_to_book",
    "ev_to_revenue", "ev_to_ebitda", "dividend_yield", "beta",
    "week_52_high", "week_52_low", "average_volume",
]


@retry()
def _fetch_info(ticker: str) -> dict:
    return get_yf_ticker(ticker).info or {}


def get_valuation_snapshot(ticker: str) -> dict:
    """Return current valuation metrics for `ticker`.

    Missing fields come back as None rather than raising -- not every
    company reports every multiple (e.g. no PEG for unprofitable
    companies, no dividend yield for non-payers).
    """
    ticker = ticker.upper()
    try:
        info = _fetch_info(ticker)
    except Exception as exc:
        logger.error("Failed to fetch valuation data for %s: %s", ticker, exc)
        return {}

    if not info:
        logger.warning("No valuation data available for %s", ticker)
        return {}

    enterprise_value = safe_get(info, "enterpriseValue")
    revenue = safe_get(info, "totalRevenue")
    ebitda = safe_get(info, "ebitda")

    ev_to_revenue = safe_get(info, "enterpriseToRevenue")
    if ev_to_revenue is None:
        ev_to_revenue = safe_divide(enterprise_value, revenue)

    ev_to_ebitda = safe_get(info, "enterpriseToEbitda")
    if ev_to_ebitda is None:
        ev_to_ebitda = safe_divide(enterprise_value, ebitda)

    return {
        "ticker": ticker,
        "as_of": datetime.today().strftime("%Y-%m-%d"),
        "price": safe_get(info, "currentPrice", "regularMarketPrice"),
        "market_cap": safe_get(info, "marketCap"),
        "enterprise_value": enterprise_value,
        "shares_outstanding": safe_get(info, "sharesOutstanding"),
        "pe_ratio": safe_get(info, "trailingPE"),
        "forward_pe": safe_get(info, "forwardPE"),
        "peg_ratio": safe_get(info, "trailingPegRatio", "pegRatio"),
        "price_to_sales": safe_get(info, "priceToSalesTrailing12Months"),
        "price_to_book": safe_get(info, "priceToBook"),
        "ev_to_revenue": ev_to_revenue,
        "ev_to_ebitda": ev_to_ebitda,
        # yfinance currently returns this as an already-scaled percentage
        # number (e.g. 2.56 meaning 2.56%), confirmed empirically against
        # known real yields (KO ~2.9%, XOM ~3.3%) -- not a decimal
        # fraction. Divide by 100 here so it matches the decimal-fraction
        # convention every other ratio in this codebase uses (0.15 =
        # 15%), which is what lets excel_workbook.py apply one uniform
        # percentage cell format everywhere instead of special-casing
        # this field.
        "dividend_yield": safe_divide(safe_get(info, "dividendYield"), 100),
        "beta": safe_get(info, "beta"),
        "week_52_high": safe_get(info, "fiftyTwoWeekHigh"),
        "week_52_low": safe_get(info, "fiftyTwoWeekLow"),
        "average_volume": safe_get(info, "averageVolume", "averageDailyVolume10Day"),
    }


def get_company_profile(ticker: str) -> dict:
    """Return descriptive company profile fields for `ticker`, used on
    the Dashboard and Watchlist tabs."""
    ticker = ticker.upper()
    try:
        info = _fetch_info(ticker)
    except Exception as exc:
        logger.error("Failed to fetch company profile for %s: %s", ticker, exc)
        return {}

    if not info:
        logger.warning("No profile data available for %s", ticker)
        return {}

    return {
        "ticker": ticker,
        "name": safe_get(info, "longName", "shortName"),
        "sector": safe_get(info, "sector"),
        "industry": safe_get(info, "industry"),
        "description": safe_get(info, "longBusinessSummary"),
        "website": safe_get(info, "website"),
        "country": safe_get(info, "country"),
        "exchange": safe_get(info, "fullExchangeName", "exchange"),
        "employees": safe_get(info, "fullTimeEmployees"),
    }
