"""
news.py

Pulls recent news for a ticker from Yahoo Finance (yfinance's
Ticker.news), normalizes each item into the fields the spec's News tab
needs, and uses ai_client.py to generate a short investor-relevant
summary, sentiment, and category for each headline.

Raw fetch/normalize is deliberately separate from AI enrichment
(get_raw_news vs get_news) so news collection still produces usable
rows -- Date, Headline, Publisher, Link, Company -- even if the AI
backend is unavailable or misconfigured. A missing API key should
degrade the News tab to "no AI fields", not break collection entirely.
"""

import logging
from datetime import datetime, timedelta, timezone
from typing import Optional

import pandas as pd

import ai_client
import tickers
from config import MAX_NEWS_ITEMS_PER_TICKER, NEWS_LOOKBACK_DAYS
from utils import extract_json_object, get_yf_ticker, retry

logger = logging.getLogger(__name__)

CATEGORIES = [
    "Earnings", "Product", "Guidance", "Acquisition", "Macro",
    "Management", "AI", "Competition", "Legal", "Regulation", "Other",
]
SENTIMENTS = ["Positive", "Negative", "Neutral"]

NEWS_COLUMNS = ["date", "headline", "publisher", "link", "company", "ticker", "ai_summary", "sentiment", "category"]


@retry()
def _fetch_raw_news(ticker: str) -> list:
    return get_yf_ticker(ticker).news or []


def _parse_item(ticker: str, item: dict) -> Optional[dict]:
    """Normalize one yfinance news item. Current schema nests everything
    under item['content']; title and pubDate are the only fields we
    treat as required -- everything else degrades to None."""
    content = item.get("content") or {}
    title = content.get("title")
    pub_date = content.get("pubDate")
    if not title or not pub_date:
        return None

    provider = (content.get("provider") or {}).get("displayName")
    url = (content.get("canonicalUrl") or {}).get("url") or (content.get("clickThroughUrl") or {}).get("url")

    return {
        "ticker": ticker,
        "date": pub_date[:10],  # ISO date portion, e.g. "2026-07-07"
        "headline": title,
        "publisher": provider,
        "link": url,
        "yahoo_summary": content.get("summary"),
    }


def get_raw_news(
    ticker: str,
    lookback_days: int = NEWS_LOOKBACK_DAYS,
    max_items: int = MAX_NEWS_ITEMS_PER_TICKER,
) -> pd.DataFrame:
    """Fetch and normalize recent news for `ticker`, no AI enrichment.
    Filtered to the last `lookback_days` and capped at `max_items`,
    newest first."""
    ticker = ticker.upper()
    try:
        raw_items = _fetch_raw_news(ticker)
    except Exception as exc:
        logger.error("Failed to fetch news for %s: %s", ticker, exc)
        return pd.DataFrame()

    cutoff = datetime.now(timezone.utc) - timedelta(days=lookback_days)
    rows = []
    for item in raw_items:
        parsed = _parse_item(ticker, item)
        if parsed is None:
            continue
        try:
            # Day-granularity comparison (parsed date has no time component) --
            # accurate enough for a "last N days" filter.
            pub_dt = datetime.fromisoformat(parsed["date"]).replace(tzinfo=timezone.utc)
        except ValueError:
            continue
        if pub_dt < cutoff:
            continue
        rows.append(parsed)

    if not rows:
        return pd.DataFrame()

    return pd.DataFrame(rows).sort_values("date", ascending=False).head(max_items).reset_index(drop=True)


def _classify_item(ticker: str, headline: str, yahoo_summary: Optional[str]) -> dict:
    """Ask the configured AI provider for a short investor-relevant
    summary, sentiment, and category for one headline.

    Returns None-valued fields on any failure (bad API key, malformed
    response, provider outage) rather than raising -- one bad headline
    or a down AI provider should never abort the whole news run.
    """
    prompt = (
        f"You are classifying a single news headline about {ticker} for an "
        f"investment research database.\n\n"
        f"Headline: {headline}\n"
        f"Yahoo summary: {yahoo_summary or '(none provided)'}\n\n"
        f"Respond with ONLY a JSON object (no markdown, no commentary) with "
        f"exactly these keys:\n"
        f'- "summary": one sentence (<=25 words) on why this matters for a '
        f"long-term investor in {ticker} -- not a restatement of the headline.\n"
        f'- "sentiment": one of {SENTIMENTS}\n'
        f'- "category": one of {CATEGORIES}\n'
    )
    try:
        raw_response = ai_client.generate(prompt, max_tokens=200)
        parsed = extract_json_object(raw_response)
        sentiment = parsed.get("sentiment")
        category = parsed.get("category")
        return {
            "ai_summary": parsed.get("summary"),
            "sentiment": sentiment if sentiment in SENTIMENTS else None,
            "category": category if category in CATEGORIES else "Other",
        }
    except Exception as exc:
        logger.warning("AI classification failed for %s headline %r: %s", ticker, headline[:60], exc)
        return {"ai_summary": None, "sentiment": None, "category": None}


def get_news(
    ticker: str,
    lookback_days: int = NEWS_LOOKBACK_DAYS,
    max_items: int = MAX_NEWS_ITEMS_PER_TICKER,
    enrich: bool = True,
) -> pd.DataFrame:
    """Full News-tab pipeline for one ticker: fetch, filter, and (unless
    enrich=False) classify each headline via the AI provider.

    enrich=False skips AI calls entirely -- useful for a fast raw-news
    check (e.g. add_company.py's validation step) or when no AI key is
    configured yet.
    """
    ticker = ticker.upper()
    df = get_raw_news(ticker, lookback_days=lookback_days, max_items=max_items)
    if df.empty:
        return df

    record = tickers.get_ticker_record(ticker)
    company_name = record["name"] if record else ticker

    if enrich:
        classified = df.apply(
            lambda r: _classify_item(ticker, r["headline"], r["yahoo_summary"]), axis=1
        )
        df = pd.concat([df.reset_index(drop=True), pd.DataFrame(list(classified))], axis=1)
    else:
        df["ai_summary"] = None
        df["sentiment"] = None
        df["category"] = None

    df["company"] = company_name
    return df[NEWS_COLUMNS]


def get_new_news(ticker: str, existing_links: set, **kwargs) -> pd.DataFrame:
    """News items for `ticker` not already present in existing_links.

    excel_workbook.py passes the set of links already written to the
    News tab; this is what keeps a daily refresh from re-adding
    duplicate rows for headlines already seen on a previous run.
    """
    df = get_news(ticker, **kwargs)
    if df.empty:
        return df
    return df[~df["link"].isin(existing_links)].reset_index(drop=True)
