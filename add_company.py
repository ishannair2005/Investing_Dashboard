"""
add_company.py

The "Add Company" workflow: validates a ticker, downloads everything
about it (profile, financials, valuation, news, an initial AI thesis),
registers it in tickers.py's universe, and writes all of it to the
Excel workbook -- in one call, with zero code changes required
afterward.

Usage:
    python add_company.py AMD
    python add_company.py AMD --no-ai          (skip the AI writeup -- faster, no API cost)
    python add_company.py AMD --dry-run         (pull and print data, skip workbook write + registration)

main.py's `--add` flag imports and calls add_company() directly rather
than shelling out to this script, so both entry points ("python
add_company.py AMD" and "python main.py --add AMD" from the spec) run
the exact same code path.

Registration (tickers.add_ticker) happens BEFORE the workbook sync, not
after: if the sync fails partway (the file was open in Excel, a bad
API key), the ticker is still in the universe and the next scheduled
run (daily/quarterly) will pick up whatever didn't get written, rather
than silently dropping the company because setup didn't 100% complete.
"""

import argparse
import logging
import sys
from typing import Optional

import tickers
import valuation
from config import IS_LOCAL_INSTANCE
from utils import get_yf_ticker, validate_ticker

logger = logging.getLogger(__name__)


def _check_price_history(ticker: str) -> Optional[str]:
    """Confirm the ticker has real trading history (the spec's "download
    historical prices" step). No dedicated Price History tab exists yet
    in the workbook structure -- daily price charting is explicitly
    listed under FUTURE EXPANSION -- so this only validates and logs a
    date range rather than persisting anything. A future prices.py
    module has a natural hook point here.
    """
    try:
        hist = get_yf_ticker(ticker).history(period="5d")
    except Exception as exc:
        logger.warning("Could not fetch price history for %s: %s", ticker, exc)
        return None
    if hist.empty:
        return None
    return f"{hist.index.min().date()} to {hist.index.max().date()}"


def add_company(ticker: str, run_ai_analysis: bool = True, sync_to_workbook: bool = True) -> dict:
    """Run the full add-company workflow for `ticker`. Returns a summary
    dict describing what happened; callers (this file's CLI, or
    main.py's --add flag) decide how to present it.
    """
    ticker = ticker.strip().upper()
    summary = {"ticker": ticker, "success": False}

    logger.info("Validating %s...", ticker)
    if not validate_ticker(ticker):
        # validate_ticker() already retries transient failures, but a
        # data-provider hiccup can still exhaust those -- don't assert
        # the ticker is definitely wrong when it might just be that.
        summary["error"] = (
            f"'{ticker}' didn't resolve to a tradeable company on Yahoo Finance. "
            f"If you're confident this ticker is correct, this may be a temporary data "
            f"provider issue (check the logs) -- try again in a moment."
        )
        return summary

    existing = tickers.get_ticker_record(ticker)
    if existing and existing.get("active", True):
        summary["error"] = (
            f"{ticker} is already tracked (added {existing.get('added_date')}). "
            f"It's already included in every automatic run -- use main.py to refresh its data."
        )
        return summary

    logger.info("Downloading company profile for %s...", ticker)
    profile = valuation.get_company_profile(ticker)
    name = profile.get("name") or ticker
    sector = profile.get("sector") or "Other"
    summary["name"] = name
    summary["sector"] = sector

    price_range = _check_price_history(ticker)
    summary["price_history_range"] = price_range
    if price_range is None:
        logger.warning("No historical price data found for %s -- proceeding anyway", ticker)

    if not sync_to_workbook:
        # --dry-run: pull and validate only. Registration is deliberately
        # skipped too (not just the workbook sync) so a dry run is truly
        # read-only and can be re-run freely without side effects.
        summary["success"] = True
        summary["workbook_synced"] = False
        return summary

    import git_sync  # deferred: keeps a --dry-run usable even if the workbook is mid-edit elsewhere
    import excel_workbook  # deferred for the same reason
    from utils import WRITE_LOCK

    # One lock held across the whole "register + sync" sequence, not a
    # separate acquisition per step. Previously each step (pull+reload,
    # register ticker, sync tabs) locked independently, which left gaps
    # where a *different* concurrent add's pull-and-reload could swap
    # out the in-memory workbook mid-sequence for this one -- splitting
    # a single ticker's writes across two different Workbook objects and
    # silently losing whatever was written to the first. This was very
    # likely the real trigger behind a production segfault, not just a
    # data-loss bug: two threads holding references to different
    # generations of a non-thread-safe openpyxl object at once.
    with WRITE_LOCK:
        if not IS_LOCAL_INSTANCE:
            # A deployed cloud instance's container may have been running
            # for a while -- pull the latest state (and drop the cached
            # in-memory workbook) before mutating, so this doesn't
            # clobber a change pushed from the local Mac or another
            # session since it started.
            git_sync.sync_before_write()

        if existing:  # inactive (previously soft-removed) -- reactivate instead of re-adding
            tickers.reactivate_ticker(ticker)
            logger.info("Reactivated %s in the universe", ticker)
        else:
            tickers.add_ticker(ticker, name=name, sector=sector)
            logger.info("Registered %s in the universe (sector=%s)", ticker, sector)

        logger.info("Initializing workbook and syncing all tabs for %s (this pulls financials, "
                    "ratios, valuation, news%s)...", ticker, ", and generates an AI thesis" if run_ai_analysis else "")
        excel_workbook.initialize_workbook()
        # Held for the full sync, including its network/AI calls, not
        # just the writes -- see the comment above the outer `with`.
        # Slower for two concurrent adds (they now fully serialize
        # instead of overlapping their fetches), but this is a personal/
        # small-team tool: an extra 30-60s wait is a much better outcome
        # than a segfault or silently losing a ticker's data.
        sync_summary = excel_workbook.sync_ticker_full(ticker, run_ai_analysis=run_ai_analysis)
        git_sync.push_state_if_changed(f"Add {ticker}")

    summary.update(sync_summary)
    summary["workbook_synced"] = True
    summary["success"] = True

    return summary


def _print_summary(summary: dict) -> None:
    ticker = summary.get("ticker", "?")
    if not summary.get("success"):
        print(f"\nFailed to add {ticker}: {summary.get('error', 'unknown error')}\n")
        return

    print(f"\nAdded {ticker} ({summary.get('name', '?')}, {summary.get('sector', '?')})")
    if summary.get("price_history_range"):
        print(f"  Price history available: {summary['price_history_range']}")

    if not summary.get("workbook_synced"):
        print("  Dry run -- workbook sync and universe registration were skipped.\n")
        return

    print(f"  New financial quarters written: {summary.get('new_financial_quarters', 0)}")
    print(f"  New ratio quarters written:     {summary.get('new_ratio_quarters', 0)}")
    print(f"  New news items written:         {summary.get('new_news_items', 0)}")
    print(f"  AI investment thesis generated: {summary.get('ai_analysis_generated', False)}")
    print(f"\n{ticker} is now part of the tracked universe -- it will be included "
          f"in every future daily/quarterly run automatically.\n")


def main() -> None:
    parser = argparse.ArgumentParser(description="Add a new company to the investment research dashboard.")
    parser.add_argument("ticker", help="Ticker symbol to add, e.g. AMD")
    parser.add_argument("--no-ai", action="store_true", help="Skip generating the initial AI investment thesis")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Pull and validate data only -- skip universe registration and workbook sync",
    )
    args = parser.parse_args()

    summary = add_company(args.ticker, run_ai_analysis=not args.no_ai, sync_to_workbook=not args.dry_run)
    _print_summary(summary)
    sys.exit(0 if summary.get("success") else 1)


if __name__ == "__main__":
    main()
