"""
ratios.py

Computes the growth, profitability, and balance-sheet ratios from the
spec's GROWTH METRICS section, using financials.py's quarterly history
as input. This module does no fetching of its own -- it's a pure
transform over the DataFrame financials.get_quarterly_financials()
returns, which keeps it trivially testable (feed it a DataFrame, get a
DataFrame back).

Two families of ratio, computed differently on purpose:

  - Point-in-time balance sheet ratios (Debt/Equity, Current Ratio,
    Quick Ratio) use a single quarter's ending balance sheet values.
  - Return ratios (ROE, ROA, ROIC, Interest Coverage) use trailing
    twelve months (TTM) income against the ending balance sheet. This
    is standard practice -- a single quarter's net income divided by
    ending equity understates annualized return by roughly 4x and
    isn't comparable to how these ratios are normally quoted.

TTM figures require four consecutive quarters of history; rows with
fewer than four prior quarters get None for TTM-dependent ratios
rather than a misleadingly partial number.
"""

import logging
from typing import Optional

import pandas as pd

from utils import pct_change, safe_divide

logger = logging.getLogger(__name__)

# Simplifying assumption for ROIC: financials.py doesn't extract pretax
# income / tax provision (not in the spec's required line items), so a
# company-specific effective tax rate isn't available. Using the flat
# US federal statutory rate is a documented approximation, not an
# attempt at each company's actual effective rate.
ASSUMED_TAX_RATE = 0.21

# Ordered list of every ratio this module computes -- the single source
# of truth for the Ratios tab's column headers in excel_workbook.py, so
# the sheet schema can't silently drift out of sync with this module.
RATIO_METRICS = [
    "gross_margin", "operating_margin", "net_margin", "fcf_margin",
    "revenue_growth_qoq", "revenue_growth_yoy", "eps_growth_yoy",
    "debt_to_equity", "current_ratio", "quick_ratio",
    "roe", "roa", "roic", "interest_coverage",
]


def _ttm_sum(series: Optional[pd.Series]) -> Optional[float]:
    """Sum of the last 4 quarters, or None if any of the 4 is missing."""
    if series is None or len(series) < 4:
        return None
    total = series.sum(min_count=4)
    return None if pd.isna(total) else float(total)


def _or_zero(value):
    return 0 if value is None or pd.isna(value) else value


def compute_ratios(df: pd.DataFrame) -> pd.DataFrame:
    """Given a quarterly financials DataFrame (as returned by
    financials.get_quarterly_financials(), any order), return a new
    DataFrame of computed ratios, one row per quarter, oldest first."""
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.sort_values("period_end").reset_index(drop=True)
    eps_col = "diluted_eps" if "diluted_eps" in df.columns and df["diluted_eps"].notna().any() else "eps"
    has_eps_col = eps_col in df.columns

    rows = []
    for i in range(len(df)):
        row = df.iloc[i]
        prior = df.iloc[i - 1] if i >= 1 else None
        year_ago = df.iloc[i - 4] if i >= 4 else None

        revenue = row.get("revenue")
        gross_profit = row.get("gross_profit")
        operating_income = row.get("operating_income")
        net_income = row.get("net_income")
        fcf = row.get("free_cash_flow")
        total_assets = row.get("total_assets")
        total_debt = row.get("total_debt")
        shareholder_equity = row.get("shareholder_equity")
        current_assets = row.get("current_assets")
        current_liabilities = row.get("current_liabilities")
        cash = row.get("cash")
        inventory = row.get("inventory")

        quick_numerator = None
        if current_assets is not None and not pd.isna(current_assets):
            quick_numerator = current_assets - _or_zero(inventory)

        ratio_row = {
            "ticker": row.get("ticker"),
            "quarter": row.get("quarter"),
            "period_end": row.get("period_end"),
            # -- Margins (single quarter) --
            "gross_margin": safe_divide(gross_profit, revenue),
            "operating_margin": safe_divide(operating_income, revenue),
            "net_margin": safe_divide(net_income, revenue),
            "fcf_margin": safe_divide(fcf, revenue),
            # -- Growth --
            "revenue_growth_qoq": pct_change(revenue, prior.get("revenue")) if prior is not None else None,
            "revenue_growth_yoy": pct_change(revenue, year_ago.get("revenue")) if year_ago is not None else None,
            "eps_growth_yoy": (
                pct_change(row.get(eps_col), year_ago.get(eps_col))
                if year_ago is not None and has_eps_col
                else None
            ),
            # -- Point-in-time balance sheet ratios --
            "debt_to_equity": safe_divide(total_debt, shareholder_equity),
            "current_ratio": safe_divide(current_assets, current_liabilities),
            "quick_ratio": safe_divide(quick_numerator, current_liabilities),
        }

        # -- TTM-based return ratios: need 4 consecutive quarters --
        if i >= 3:
            ttm_slice = df.iloc[i - 3 : i + 1]
            ttm_net_income = _ttm_sum(ttm_slice["net_income"]) if "net_income" in ttm_slice.columns else None
            ttm_ebit = _ttm_sum(ttm_slice["ebit"]) if "ebit" in ttm_slice.columns else None
            ttm_interest_expense = (
                _ttm_sum(ttm_slice["interest_expense"]) if "interest_expense" in ttm_slice.columns else None
            )

            invested_capital = None
            if total_debt is not None and shareholder_equity is not None:
                invested_capital = total_debt + shareholder_equity - _or_zero(cash)

            nopat = ttm_ebit * (1 - ASSUMED_TAX_RATE) if ttm_ebit is not None else None

            ratio_row["roe"] = safe_divide(ttm_net_income, shareholder_equity)
            ratio_row["roa"] = safe_divide(ttm_net_income, total_assets)
            ratio_row["roic"] = safe_divide(nopat, invested_capital)
            ratio_row["interest_coverage"] = (
                safe_divide(ttm_ebit, abs(ttm_interest_expense)) if ttm_interest_expense else None
            )
        else:
            ratio_row["roe"] = None
            ratio_row["roa"] = None
            ratio_row["roic"] = None
            ratio_row["interest_coverage"] = None

        rows.append(ratio_row)

    return pd.DataFrame(rows)


def get_ratios(ticker: str) -> pd.DataFrame:
    """Convenience wrapper: fetch a ticker's quarterly financials and
    compute ratios in one call."""
    import financials  # deferred import: avoids a hard import-time cycle

    df = financials.get_quarterly_financials(ticker)
    return compute_ratios(df)


def get_latest_ratios(ticker: str) -> dict:
    """Most recent quarter's ratios as a flat dict, for the Dashboard
    tab and investment scoring."""
    df = get_ratios(ticker)
    if df.empty:
        return {}
    return df.iloc[-1].to_dict()
