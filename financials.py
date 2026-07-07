"""
financials.py

Pulls quarterly financial statements (income statement, balance sheet,
cash flow) from Yahoo Finance via yfinance and normalizes them into one
row per quarter, keyed by utils.quarter_label() (e.g. "Q2 2026").

This module never talks to the Excel workbook directly -- it only
knows how to fetch and normalize. It exposes two entry points:

  - get_quarterly_financials(ticker): the full normalized history
    currently available from yfinance.
  - get_new_quarters(ticker, existing_labels): the subset of quarters
    NOT already in existing_labels.

excel_workbook.py (a later module) is what actually knows which quarters
are already stored, by reading the Financials tab. It will call
get_new_quarters() with that set and append only what comes back --
that's the mechanism behind "append new quarters, never overwrite
history."
"""

import logging
from typing import Optional

import pandas as pd

from utils import get_yf_ticker, quarter_label, retry

logger = logging.getLogger(__name__)

# Each metric maps to a list of candidate yfinance row labels, tried in
# order. yfinance's statement row names have shifted across library
# versions and occasionally differ by company/filing type, so we don't
# rely on a single exact name.
_INCOME_ROWS = {
    "revenue": ["Total Revenue", "Operating Revenue"],
    "gross_profit": ["Gross Profit"],
    "operating_income": ["Operating Income"],
    "ebit": ["EBIT"],
    "ebitda": ["EBITDA", "Normalized EBITDA"],
    "net_income": ["Net Income", "Net Income Common Stockholders"],
    "eps": ["Basic EPS"],
    "diluted_eps": ["Diluted EPS"],
    "interest_expense": ["Interest Expense", "Interest Expense Non Operating"],
}

_BALANCE_ROWS = {
    "cash": ["Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments"],
    "total_assets": ["Total Assets"],
    "current_assets": ["Current Assets"],
    "total_debt": ["Total Debt"],
    "long_term_debt": ["Long Term Debt", "Long Term Debt And Capital Lease Obligation"],
    "shareholder_equity": ["Stockholders Equity", "Total Equity Gross Minority Interest"],
    "current_liabilities": ["Current Liabilities"],
    "inventory": ["Inventory"],
}

_CASHFLOW_ROWS = {
    "operating_cash_flow": ["Operating Cash Flow", "Cash Flow From Continuing Operating Activities"],
    "capital_expenditures": ["Capital Expenditure"],
    "free_cash_flow": ["Free Cash Flow"],
    "investing_cash_flow": ["Investing Cash Flow", "Cash Flow From Continuing Investing Activities"],
    "financing_cash_flow": ["Financing Cash Flow", "Cash Flow From Continuing Financing Activities"],
}

ALL_METRICS = list(_INCOME_ROWS) + list(_BALANCE_ROWS) + list(_CASHFLOW_ROWS)


def _find_row(df: pd.DataFrame, candidates: list) -> Optional[pd.Series]:
    """Return the first matching row (indexed by period-end date) out of
    a list of candidate labels, or None if none of them exist in df."""
    if df is None or df.empty:
        return None
    for label in candidates:
        if label in df.index:
            return df.loc[label]
    return None


def _extract(df: pd.DataFrame, row_map: dict) -> dict:
    """Build {metric_name: {period_end: value}} for every metric in
    row_map that has a matching row in df."""
    result = {}
    for metric, candidates in row_map.items():
        row = _find_row(df, candidates)
        if row is not None:
            result[metric] = row.to_dict()
    return result


@retry()
def _fetch_statements(ticker: str):
    yf_ticker = get_yf_ticker(ticker)
    return (
        yf_ticker.quarterly_income_stmt,
        yf_ticker.quarterly_balance_sheet,
        yf_ticker.quarterly_cashflow,
    )


def get_quarterly_financials(ticker: str) -> pd.DataFrame:
    """Return one row per available quarter for `ticker`.

    Columns: ticker, quarter, period_end, plus every metric in
    ALL_METRICS that yfinance provided data for. Rows are sorted oldest
    -> newest so this reads naturally as a time series and appends to a
    sheet in chronological order.
    """
    ticker = ticker.upper()
    try:
        income, balance, cashflow = _fetch_statements(ticker)
    except Exception as exc:
        logger.error("Failed to fetch financial statements for %s: %s", ticker, exc)
        return pd.DataFrame()

    income_data = _extract(income, _INCOME_ROWS)
    balance_data = _extract(balance, _BALANCE_ROWS)
    cashflow_data = _extract(cashflow, _CASHFLOW_ROWS)

    # Union of every period-end date seen across all three statements --
    # they don't always share an identical set of columns.
    period_ends = set()
    for metric_data in (income_data, balance_data, cashflow_data):
        for values in metric_data.values():
            period_ends.update(values.keys())

    if not period_ends:
        logger.warning("No quarterly financial data available for %s", ticker)
        return pd.DataFrame()

    rows = []
    for period_end in sorted(period_ends):
        row = {
            "ticker": ticker,
            "quarter": quarter_label(period_end),
            "period_end": pd.Timestamp(period_end).date().isoformat(),
        }
        for metric_data in (income_data, balance_data, cashflow_data):
            for metric, values in metric_data.items():
                row[metric] = values.get(period_end)
        rows.append(row)

    df = pd.DataFrame(rows)

    # Free Cash Flow: use yfinance's own figure if present, otherwise
    # derive it. yfinance reports Capital Expenditure as a negative
    # (cash outflow), so OCF + CapEx = OCF - |CapEx|.
    if "free_cash_flow" not in df.columns:
        df["free_cash_flow"] = None
    needs_fcf = df["free_cash_flow"].isna()
    if needs_fcf.any() and "operating_cash_flow" in df.columns and "capital_expenditures" in df.columns:
        df.loc[needs_fcf, "free_cash_flow"] = (
            df.loc[needs_fcf, "operating_cash_flow"] + df.loc[needs_fcf, "capital_expenditures"]
        )

    ordered_columns = ["ticker", "quarter", "period_end"] + [m for m in ALL_METRICS if m in df.columns]
    df = df[ordered_columns].sort_values("period_end").reset_index(drop=True)
    return df


def get_new_quarters(ticker: str, existing_labels: set, df: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """Quarters for `ticker` not already present in existing_labels.

    excel_workbook.py passes the set of quarter labels already written
    to the Financials tab for this ticker; every row this returns is
    safe to append.

    Accepts an already-fetched `df` (get_quarterly_financials' output)
    so callers that need the full history for another purpose too
    (e.g. sync_ticker_full computing ratios in the same run) don't
    trigger a second yfinance round-trip for the same ticker.
    """
    if df is None:
        df = get_quarterly_financials(ticker)
    if df.empty:
        return df
    return df[~df["quarter"].isin(existing_labels)].reset_index(drop=True)
