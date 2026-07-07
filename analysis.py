"""
analysis.py

Produces the full 10-part independent investment writeup described in
the spec's AI INVESTMENT ANALYSIS section: executive summary, key
developments, bull/bear case, financial health, valuation assessment,
competitive position, long-term outlook, investment thesis, and a
confidence score.

Design principles carried over from the spec's REASONING REQUIREMENTS
directly into the system prompt below: no reliance on analyst ratings
or price targets, independent cross-referenced reasoning, explicit
fact-vs-inference distinction, stated uncertainty on mixed evidence,
and valuation judged against the company's own historical range rather
than a bare P/E comparison.

This module gathers evidence from financials.py, ratios.py,
valuation.py, and news.py, hands it to ai_client.py as structured JSON
(not freeform prose) so the model reasons over real numbers instead of
inventing them, and parses the response back into a dict. It never
talks to the Excel workbook -- excel_workbook.py supplies
`previous_analysis` (read from the Narrative tab's last row for this
ticker) and persists whatever comes back.
"""

import json
import logging
from datetime import datetime
from typing import Optional

import pandas as pd

import financials
import news
import ratios
import tickers
import valuation
import ai_client
from config import NEWS_LOOKBACK_DAYS
from utils import extract_json_object

logger = logging.getLogger(__name__)

# Ordered list of every field an analysis produces -- the single source
# of truth for both response validation (REQUIRED_KEYS) and the
# Narrative tab's column order in excel_workbook.py, so the sheet schema
# can't silently drift out of sync with this module.
NARRATIVE_FIELDS = [
    "executive_summary",
    "key_developments",
    "bull_case",
    "bear_case",
    "financial_health",
    "valuation_rating",
    "valuation_assessment",
    "competitive_position",
    "long_term_outlook",
    "investment_thesis",
    "confidence_score",
    "confidence_rationale",
    "competitive_position_score",
    "long_term_outlook_score",
]
REQUIRED_KEYS = set(NARRATIVE_FIELDS)
VALUATION_RATINGS = {"Undervalued", "Fairly Valued", "Overvalued"}

# Trailing window of quarters sent to the model: long enough to show a
# multi-year trend in margins/returns, short enough to keep the prompt
# (and cost) bounded.
TRAILING_QUARTERS = 8

_SYSTEM_PROMPT = """You are a senior long-term equity analyst at a hedge fund, performing independent investment due diligence. You reason like a portfolio manager conducting due diligence, not a news summarizer.

Rules you must follow:
- Do not rely on or cite analyst ratings, consensus price targets, or media sentiment as evidence. Reach your own conclusions strictly from the financial data and news evidence provided.
- Do not copy language from news headlines or summaries -- synthesize what they imply, don't quote them.
- Cross-reference multiple pieces of evidence before drawing a conclusion; when two pieces of evidence conflict, explain both sides before reaching a view.
- Explicitly distinguish facts (directly grounded in the data provided) from inferences (your reasoning about what those facts imply).
- State uncertainty clearly when evidence is mixed or incomplete. Do not manufacture false confidence.
- Prioritize long-term business fundamentals -- multi-year trends in margins, returns on capital, competitive position -- over short-term stock price movements.
- Judge valuation against the company's own historical multiple range and business quality, not a bare "P/E is high therefore overvalued" heuristic.
- Only describe something as a "key development" if it represents a material change from the previous analysis provided. If nothing material changed, say so plainly instead of manufacturing new content.
"""


def build_evidence(
    ticker: str,
    financials_df: Optional[pd.DataFrame] = None,
    news_df: Optional[pd.DataFrame] = None,
) -> dict:
    """Gather everything the model needs to reason about `ticker` as
    structured data, not prose -- this is what keeps the analysis
    grounded in real figures instead of inviting hallucination.

    Accepts pre-fetched financials_df / news_df so callers that already
    pulled this data for another tab this run (excel_workbook.py's
    sync_ticker_full) don't trigger duplicate yfinance calls or, more
    importantly, duplicate paid AI classification calls for the same
    headlines.
    """
    profile = valuation.get_company_profile(ticker)
    val_snapshot = valuation.get_valuation_snapshot(ticker)

    fin_df = financials_df if financials_df is not None else financials.get_quarterly_financials(ticker)
    ratios_df = ratios.compute_ratios(fin_df)
    if news_df is None:
        news_df = news.get_news(ticker, enrich=True)

    sector = profile.get("sector") or ""
    sector_peers = [r["ticker"] for r in tickers.get_tickers_by_sector(sector) if r["ticker"] != ticker]

    return {
        "ticker": ticker,
        "profile": profile,
        "valuation": val_snapshot,
        "trailing_financials": fin_df.tail(TRAILING_QUARTERS).to_dict(orient="records") if not fin_df.empty else [],
        "trailing_ratios": ratios_df.tail(TRAILING_QUARTERS).to_dict(orient="records") if not ratios_df.empty else [],
        "recent_news": news_df.to_dict(orient="records") if not news_df.empty else [],
        "sector_peers": sector_peers,
    }


def _build_prompt(evidence: dict, previous_analysis: Optional[dict]) -> str:
    previous_block = "No previous analysis exists for this company -- this is the first writeup."
    if previous_analysis:
        previous_block = (
            "PREVIOUS ANALYSIS (use this to identify what has materially changed):\n"
            f"{json.dumps(previous_analysis, indent=2, default=str)}"
        )

    return f"""Analyze {evidence['ticker']} using only the evidence below. All monetary figures are USD unless noted.

COMPANY PROFILE:
{json.dumps(evidence['profile'], indent=2, default=str)}

CURRENT VALUATION SNAPSHOT:
{json.dumps(evidence['valuation'], indent=2, default=str)}

TRAILING QUARTERLY FINANCIALS (oldest to newest, up to {TRAILING_QUARTERS} quarters):
{json.dumps(evidence['trailing_financials'], indent=2, default=str)}

TRAILING QUARTERLY RATIOS (oldest to newest, up to {TRAILING_QUARTERS} quarters):
{json.dumps(evidence['trailing_ratios'], indent=2, default=str)}

RECENT NEWS (last {NEWS_LOOKBACK_DAYS} days, with sentiment/category already classified):
{json.dumps(evidence['recent_news'], indent=2, default=str)}

SECTOR PEERS (for competitive-position context; no financial data provided for them): {evidence['sector_peers']}

{previous_block}

Respond with ONLY a JSON object (no markdown fences, no commentary outside the JSON) with exactly these keys:

{{
  "executive_summary": "5-10 sentences on the company's current situation",
  "key_developments": "what changed since the previous analysis and why it matters; if no previous analysis, the most important recent developments",
  "bull_case": "strongest reasons this could outperform over the next 3-5 years",
  "bear_case": "strongest risks and reasons the investment thesis could fail",
  "financial_health": "revenue quality, margin trends, cash generation, balance sheet strength, debt sustainability, capital allocation, dilution/buybacks, return on capital",
  "valuation_rating": "one of: Undervalued, Fairly Valued, Overvalued",
  "valuation_assessment": "reasoning for the rating using multiples and historical comparison, not just the current P/E",
  "competitive_position": "economic moat, market share, industry structure, competitive threats, switching costs, pricing power",
  "long_term_outlook": "growth drivers, risks, industry tailwinds/headwinds, expected challenges over 5 years",
  "investment_thesis": "evidence-based thesis synthesizing the sections above",
  "confidence_score": "integer 0-100",
  "confidence_rationale": "what specifically increases or decreases your confidence",
  "competitive_position_score": "integer 0-100 rating the strength of the moat/competitive position you described above (0=no moat/commodity business, 100=dominant, durable moat)",
  "long_term_outlook_score": "integer 0-100 rating the long-term outlook you described above (0=structurally declining, 100=exceptionally strong multi-year growth runway)"
}}
"""


def generate_analysis(
    ticker: str,
    previous_analysis: Optional[dict] = None,
    evidence: Optional[dict] = None,
) -> Optional[dict]:
    """Produce the full investment writeup for `ticker`.

    Returns None (logged) if evidence can't be gathered or the AI call
    fails/returns an unparseable or incomplete response. Callers
    (main.py) should skip the ticker and continue to the next one
    rather than aborting the whole run -- consistent with every other
    collection module in this project.

    Accepts pre-built `evidence` (from build_evidence()) so callers that
    already assembled it for another purpose this run don't pay for a
    second evidence-gathering pass.
    """
    ticker = ticker.upper()
    if evidence is None:
        try:
            evidence = build_evidence(ticker)
        except Exception as exc:
            logger.error("Failed to gather evidence for %s: %s", ticker, exc)
            return None

    if not evidence["trailing_financials"] and not evidence["valuation"]:
        logger.warning("Insufficient data to analyze %s -- skipping", ticker)
        return None

    prompt = _build_prompt(evidence, previous_analysis)
    try:
        raw_response = ai_client.generate(prompt, system=_SYSTEM_PROMPT, max_tokens=4096)
        result = extract_json_object(raw_response)
    except Exception as exc:
        logger.error("AI analysis failed for %s: %s", ticker, exc)
        return None

    missing = REQUIRED_KEYS - result.keys()
    if missing:
        logger.error("AI analysis for %s missing required keys: %s", ticker, missing)
        return None

    if result.get("valuation_rating") not in VALUATION_RATINGS:
        logger.warning(
            "AI returned invalid valuation_rating %r for %s -- defaulting to 'Fairly Valued'",
            result.get("valuation_rating"), ticker,
        )
        result["valuation_rating"] = "Fairly Valued"

    for score_field in ("confidence_score", "competitive_position_score", "long_term_outlook_score"):
        try:
            result[score_field] = max(0, min(100, int(result[score_field])))
        except (TypeError, ValueError):
            logger.warning("AI returned non-integer %s for %s: %r", score_field, ticker, result.get(score_field))
            result[score_field] = None

    result["ticker"] = ticker
    result["generated_at"] = datetime.today().strftime("%Y-%m-%d")
    return result
