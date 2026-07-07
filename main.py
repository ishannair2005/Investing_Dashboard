"""
main.py

CLI entry point tying every module together into the automation the
spec describes:

    python main.py --daily                 refresh news + AI analysis for every tracked ticker
    python main.py --quarterly              refresh financials + ratios (+ everything else) for every tracked ticker
    python main.py --add AMD                identical to `python add_company.py AMD`
    python main.py --daily --ticker AAPL    limit a run to one ticker instead of the whole universe
    python main.py --quarterly --no-ai      skip AI analysis generation (faster, no API cost)

Every ticker is processed independently, and one ticker's failure never
aborts the run for the rest -- the spec's "if one ticker fails, continue
processing the rest" requirement is enforced here, at the top of the
call stack, not just inside individual fetch functions. A scheduler
(cron, GitHub Actions) is expected to call --daily once a day and
--quarterly once a quarter (or after detecting an earnings release).
"""

import argparse
import logging
import sys
from typing import Optional

import excel_workbook
import tickers
from add_company import add_company, _print_summary

logger = logging.getLogger(__name__)


def _run_for_tickers(ticker_list: list, run_financials: bool, run_ai_analysis: bool) -> dict:
    excel_workbook.initialize_workbook()

    results = {"succeeded": [], "failed": []}
    total = len(ticker_list)
    for i, ticker in enumerate(ticker_list, start=1):
        logger.info("[%d/%d] Syncing %s...", i, total, ticker)
        try:
            summary = excel_workbook.sync_ticker_full(
                ticker, run_financials=run_financials, run_ai_analysis=run_ai_analysis
            )
            results["succeeded"].append(summary)
            logger.info(
                "[%d/%d] %s done -- new_quarters=%s new_news=%s",
                i, total,
                ticker, summary.get("new_financial_quarters"), summary.get("new_news_items"),
            )
        except Exception as exc:
            logger.error("[%d/%d] %s failed: %s", i, total, ticker, exc)
            results["failed"].append({"ticker": ticker, "error": str(exc)})

    return results


def run_daily(ticker_list: Optional[list] = None, run_ai_analysis: bool = True) -> dict:
    """Daily automation: refresh news + AI analysis for every tracked
    ticker. Financials/ratios are skipped -- new quarters appear a
    handful of times a year, so checking for them daily across every
    ticker is pure wasted Sheets API quota; that's the quarterly job.
    """
    ticker_list = ticker_list or tickers.get_all_tickers()
    logger.info("Starting daily run for %d ticker(s)", len(ticker_list))
    return _run_for_tickers(ticker_list, run_financials=False, run_ai_analysis=run_ai_analysis)


def run_quarterly(ticker_list: Optional[list] = None, run_ai_analysis: bool = True) -> dict:
    """Quarterly automation: detect and append new earnings, update
    ratios, and refresh valuation/news/AI analysis/dashboards for every
    tracked ticker -- the full sync.
    """
    ticker_list = ticker_list or tickers.get_all_tickers()
    logger.info("Starting quarterly run for %d ticker(s)", len(ticker_list))
    return _run_for_tickers(ticker_list, run_financials=True, run_ai_analysis=run_ai_analysis)


def _print_run_summary(mode: str, results: dict) -> None:
    succeeded, failed = results["succeeded"], results["failed"]
    print(f"\n{mode} run complete: {len(succeeded)} succeeded, {len(failed)} failed")

    if failed:
        print("Failed tickers:")
        for f in failed:
            print(f"  {f['ticker']}: {f['error']}")
    print()


def main() -> None:
    parser = argparse.ArgumentParser(description="Investment Research Dashboard automation.")
    mode_group = parser.add_mutually_exclusive_group(required=True)
    mode_group.add_argument("--daily", action="store_true", help="Run the daily automation (news + AI analysis)")
    mode_group.add_argument(
        "--quarterly", action="store_true",
        help="Run the quarterly automation (financials + ratios + everything else)",
    )
    mode_group.add_argument("--add", metavar="TICKER", help="Add a new company (same as add_company.py TICKER)")
    parser.add_argument(
        "--ticker", metavar="TICKER",
        help="Limit --daily/--quarterly to a single ticker instead of the whole universe",
    )
    parser.add_argument(
        "--no-ai", action="store_true",
        help="Skip AI analysis generation (applies to --daily, --quarterly, and --add)",
    )
    args = parser.parse_args()

    if args.add:
        summary = add_company(args.add, run_ai_analysis=not args.no_ai)
        _print_summary(summary)
        sys.exit(0 if summary.get("success") else 1)

    ticker_list = [args.ticker.strip().upper()] if args.ticker else None

    try:
        if args.daily:
            results = run_daily(ticker_list, run_ai_analysis=not args.no_ai)
            _print_run_summary("Daily", results)
        else:  # args.quarterly
            results = run_quarterly(ticker_list, run_ai_analysis=not args.no_ai)
            _print_run_summary("Quarterly", results)
    except (RuntimeError, FileNotFoundError) as exc:
        # Setup-level failures (missing credentials.json, sheet not
        # shared with the service account) happen before any per-ticker
        # try/except gets a chance to run -- surface these as a clean,
        # actionable message instead of a raw traceback, since this is
        # the most likely first-run experience.
        print(f"\nCould not start the sync: {exc}\n")
        sys.exit(1)

    # A scheduler should treat "some tickers failed" as a warning, not a
    # fatal run -- per-ticker resilience is the whole point. Exit 1 only
    # when nothing succeeded at all, which usually means a systemic
    # problem (bad credentials, no network) rather than one flaky ticker.
    sys.exit(1 if results["failed"] and not results["succeeded"] else 0)


if __name__ == "__main__":
    main()
