"""
dashboard_app.py

Interactive local dashboard (Streamlit) for browsing the Investment
Research workbook: switch between tracked companies, review financials/
ratios/valuation/news/AI theses, add or remove companies from a search
box, edit Watchlist notes, and drop into the raw sheet tables when you
want to see exactly what's stored.

This is a VIEW layer, not a second data store: everything it displays
is read straight from Investment_Research_Dashboard.xlsx (the same file
main.py's daily/quarterly automation writes to), and the only writes it
performs are (a) registering/removing a ticker via tickers.py +
add_company.py -- the same functions the CLI uses -- and (b) saving
Watchlist manual-field edits via excel_workbook.update_watchlist_manual_fields.
No business logic is duplicated here.

Doubles as the source for the Streamlit Community Cloud deployment,
which reads this same repo's workbook snapshot. That instance sets
IS_LOCAL_INSTANCE=false in its Secrets. Write actions (add/remove
company, edit Watchlist) are still available there as long as
GITHUB_TOKEN is also set (config.CAN_EDIT_REMOTELY) -- the cloud
instance's filesystem is ephemeral and has no cached git credentials of
its own, so git_sync.py pushes on its behalf using that token instead.
Without a token, write actions are hidden and it's read-only, since an
edit made there would otherwise silently vanish on next redeploy rather
than actually persisting.

Run with:
    streamlit run dashboard_app.py
"""

import subprocess
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st

import add_company
import excel_workbook
import git_sync
import tickers
from config import CAN_EDIT_REMOTELY, EXCEL_FILE_PATH, IS_LOCAL_INSTANCE

st.set_page_config(page_title="Investment Research Dashboard", page_icon="📊", layout="wide")

# --------------------------------------------------------------------------
# Pastel styling -- the .streamlit/config.toml sets the base theme colors;
# this adds card/badge/button polish the base theme can't express.
# --------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; }
    div[data-testid="stMetric"] {
        background: #F8F2F6;
        border-radius: 12px;
        padding: 12px 16px;
        border: 1px solid #EFE1EC;
    }
    .news-card {
        background: #FBF8FC;
        border: 1px solid #F0E4EC;
        border-radius: 12px;
        padding: 14px 18px;
        margin-bottom: 12px;
    }
    .news-card a { color: #4A4453; text-decoration: none; font-weight: 700; font-size: 1.02em; }
    .news-card a:hover { text-decoration: underline; }
    .news-summary { color: #6b6377; font-size: 0.92em; margin: 6px 0; }
    .news-meta { color: #a89fb3; font-size: 0.8em; }
    .badge {
        display: inline-block; padding: 3px 11px; border-radius: 12px;
        font-size: 0.8em; font-weight: 600; margin-right: 6px;
    }
    .company-desc { color: #6b6377; font-size: 0.95em; line-height: 1.5; }
    </style>
    """,
    unsafe_allow_html=True,
)

RATING_COLORS = {
    "Undervalued": ("#D7F0DB", "#2E7D4F"),
    "Overvalued": ("#FBE0E0", "#B14848"),
    "Fairly Valued": ("#FDF1D6", "#9C7A1E"),
}
SENTIMENT_COLORS = {
    "Positive": ("#D7F0DB", "#2E7D4F"),
    "Negative": ("#FBE0E0", "#B14848"),
    "Neutral": ("#E6E6F0", "#5B5B7A"),
}
CATEGORY_COLOR = ("#E9E3F5", "#5B4B8A")


def badge(text: str, colors: tuple) -> str:
    bg, fg = colors
    return f'<span class="badge" style="background:{bg};color:{fg};">{text}</span>'


# --------------------------------------------------------------------------
# Data loading -- cached and keyed on the workbook's mtime, so edits made
# by this app or by a background daily/quarterly sync are picked up on
# the next rerun without a stale cache lingering.
# --------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def _read_workbook(mtime: float) -> dict:
    if not EXCEL_FILE_PATH.exists():
        return {}
    raw = pd.read_excel(EXCEL_FILE_PATH, sheet_name=None, engine="openpyxl")
    # pandas takes column names from row 1's literal (humanized) text --
    # map them back to the snake_case field names this app's logic and
    # excel_workbook.format_value_for_display() key off of.
    return {name: excel_workbook.dehumanize_columns(df, name) for name, df in raw.items()}


def load_workbook() -> dict:
    mtime = EXCEL_FILE_PATH.stat().st_mtime if EXCEL_FILE_PATH.exists() else 0.0
    return _read_workbook(mtime)


def sheet_for_ticker(sheets: dict, tab_name: str, ticker: str) -> pd.DataFrame:
    df = sheets.get(tab_name)
    if df is None or df.empty or "ticker" not in df.columns:
        return pd.DataFrame()
    return df[df["ticker"] == ticker].copy()


def latest_row(df: pd.DataFrame, sort_col: str) -> Optional[dict]:
    if df is None or df.empty or sort_col not in df.columns:
        return None
    return df.sort_values(sort_col).iloc[-1].to_dict()


def fmt(header: str, value) -> str:
    return excel_workbook.format_value_for_display(header, value)


def escape_markdown_math(text) -> str:
    """Streamlit's markdown renderer treats $...$ as inline LaTeX math.
    Financial text is full of dollar figures ("$8.2B", "$90B"), so any
    AI-generated writeup or news headline/summary rendered through
    st.write()/st.markdown() gets silently mangled into garbled italic
    math notation unless every literal $ is escaped first.
    """
    if text is None:
        return ""
    return str(text).replace("$", "\\$")


def fmt_compact_currency(value) -> str:
    """Abbreviated $ notation (e.g. "$4.61T") for tight metric-card
    layouts, where format_value_for_display's fully comma-grouped
    figure (e.g. "$4,607,423,545,344") would overflow the card."""
    if value is None or pd.isna(value):
        return "—"
    abs_value = abs(value)
    for threshold, suffix in ((1e12, "T"), (1e9, "B"), (1e6, "M"), (1e3, "K")):
        if abs_value >= threshold:
            return f"${value / threshold:,.2f}{suffix}"
    return f"${value:,.2f}"


# --------------------------------------------------------------------------
# Sidebar: add-company search box, tracked-company list (grouped by
# sector, with remove buttons), workbook access, view mode toggle.
# --------------------------------------------------------------------------

def render_sidebar() -> None:
    st.sidebar.markdown("## 📊 Investment Dashboard")

    if not IS_LOCAL_INSTANCE:
        if CAN_EDIT_REMOTELY:
            st.sidebar.info(
                "🌐 **Cloud instance** -- changes here are committed and pushed to GitHub, which "
                "briefly reloads this app to pick them up."
            )
        else:
            st.sidebar.info(
                "🌐 **Cloud view (read-only)** -- showing the workbook as of the last sync from the "
                "local Mac. Add/remove companies and Watchlist edits are only available there."
            )

    if CAN_EDIT_REMOTELY:
        st.sidebar.markdown("### Add a company")
        with st.sidebar.form("add_company_form", clear_on_submit=True):
            new_ticker = st.text_input("Ticker symbol", placeholder="e.g. AMD, TSM, SPOT").strip().upper()
            run_ai = st.checkbox("Generate AI thesis now", value=True, help="Slower and uses an API call, but the full writeup is ready immediately instead of waiting for the next automated sync.")
            submitted = st.form_submit_button("Add Company", width="stretch")

        if submitted:
            if not new_ticker:
                st.sidebar.warning("Enter a ticker symbol first.")
            else:
                with st.spinner(f"Adding {new_ticker} -- fetching financials, valuation, news{', running AI analysis' if run_ai else ''}..."):
                    if not IS_LOCAL_INSTANCE:
                        git_sync.sync_before_write()
                    result = add_company.add_company(new_ticker, run_ai_analysis=run_ai)
                if result.get("success"):
                    st.sidebar.success(f"Added {new_ticker}")
                    st.cache_data.clear()
                    st.session_state.selected_ticker = new_ticker
                    st.rerun()
                else:
                    st.sidebar.error(result.get("error", "Failed to add company"))

    st.sidebar.markdown("### Companies")
    records = tickers.get_ticker_records(active_only=True)
    if not records:
        st.sidebar.caption("No companies tracked yet -- add one above.")
    else:
        by_sector: dict = {}
        for r in sorted(records, key=lambda r: (r["sector"], r["ticker"])):
            by_sector.setdefault(r["sector"], []).append(r)

        selected = st.session_state.get("selected_ticker")
        for sector, recs in sorted(by_sector.items()):
            with st.sidebar.expander(f"{sector} ({len(recs)})", expanded=(selected in [r["ticker"] for r in recs] if selected else False)):
                for r in recs:
                    t = r["ticker"]
                    label = f"**{t}**" if t == selected else t
                    if CAN_EDIT_REMOTELY:
                        c1, c2 = st.columns([4, 1])
                    else:
                        c1 = st.container()
                    if c1.button(label, key=f"select_{t}", width="stretch"):
                        st.session_state.selected_ticker = t
                        st.session_state.view_mode = "company"
                        st.rerun()
                    if CAN_EDIT_REMOTELY and c2.button("🗑", key=f"remove_{t}", help=f"Remove {t} from the tracked universe"):
                        if not IS_LOCAL_INSTANCE:
                            git_sync.sync_before_write()
                        tickers.remove_ticker(t)
                        git_sync.push_state_if_changed(f"Remove {t}")
                        if st.session_state.get("selected_ticker") == t:
                            st.session_state.selected_ticker = None
                        st.cache_data.clear()
                        st.rerun()

    st.sidebar.divider()
    st.sidebar.markdown("### Workbook")
    if IS_LOCAL_INSTANCE and st.sidebar.button("📂 Open in Excel", width="stretch"):
        subprocess.run(["open", str(EXCEL_FILE_PATH)])
    if st.sidebar.button("🗂 Browse raw sheets", width="stretch"):
        st.session_state.view_mode = "raw"
        st.rerun()
    if EXCEL_FILE_PATH.exists():
        mtime = datetime.fromtimestamp(EXCEL_FILE_PATH.stat().st_mtime)
        st.sidebar.caption(f"Workbook last updated: {mtime.strftime('%Y-%m-%d %H:%M')}")


# --------------------------------------------------------------------------
# Raw sheets viewer -- "access the sheets from the app" without leaving it.
# --------------------------------------------------------------------------

def render_raw_sheets(sheets: dict) -> None:
    st.markdown("## Raw Sheets")
    if not sheets:
        st.info("No workbook found yet -- add a company or run a sync first.")
        return
    tab_name = st.selectbox("Sheet", list(sheets.keys()))
    df = sheets.get(tab_name, pd.DataFrame())
    st.caption(f"{len(df)} rows x {len(df.columns)} columns")
    st.dataframe(df, width="stretch", height=650)


# --------------------------------------------------------------------------
# Company view -- Overview / Financials / Ratios / Valuation / News /
# AI Analysis / Watchlist tabs for the selected ticker.
# --------------------------------------------------------------------------

def render_overview_tab(sheets: dict, ticker: str) -> None:
    dash_row = latest_row(sheet_for_ticker(sheets, "Dashboard", ticker), "ticker")
    if not dash_row:
        st.info(f"No Dashboard data for {ticker} yet -- it may not have synced.")
        return

    st.markdown(f"# {dash_row.get('name') or ticker} ({ticker})")
    sector_badge = badge(dash_row.get("sector") or "—", CATEGORY_COLOR)
    industry_badge = badge(dash_row.get("industry") or "—", CATEGORY_COLOR)
    st.markdown(sector_badge + industry_badge, unsafe_allow_html=True)

    if dash_row.get("description"):
        st.markdown(f'<p class="company-desc">{escape_markdown_math(dash_row["description"])}</p>', unsafe_allow_html=True)

    st.write("")
    cols = st.columns(6)
    cols[0].metric("Price", fmt("price", dash_row.get("price")))
    cols[1].metric("Market Cap", fmt_compact_currency(dash_row.get("market_cap")))
    cols[2].metric("P/E Ratio", fmt("pe_ratio", dash_row.get("pe_ratio")))
    cols[3].metric("Forward P/E", fmt("forward_pe", dash_row.get("forward_pe")))
    cols[4].metric("PEG Ratio", fmt("peg_ratio", dash_row.get("peg_ratio")))
    cols[5].metric("Dividend Yield", fmt("dividend_yield", dash_row.get("dividend_yield")))

    cols2 = st.columns(4)
    cols2[0].metric("Revenue Growth YoY", fmt("revenue_growth_yoy", dash_row.get("revenue_growth_yoy")))
    cols2[1].metric("Net Margin", fmt("net_margin", dash_row.get("net_margin")))
    cols2[2].metric("ROE", fmt("roe", dash_row.get("roe")))
    cols2[3].metric("Debt / Equity", fmt("debt_to_equity", dash_row.get("debt_to_equity")))

    st.write("")
    rating = dash_row.get("valuation_rating")
    if rating and pd.notna(rating):
        st.markdown(f"**Valuation:** {badge(rating, RATING_COLORS.get(rating, CATEGORY_COLOR))}", unsafe_allow_html=True)

    # Pull the full thesis from Narrative rather than Dashboard's
    # "latest_recommendation" copy, which is deliberately hard-truncated
    # to 200 characters to fit an Excel cell -- fine for the spreadsheet,
    # but reads as abruptly cut off here where there's room to show it in full.
    narrative_row = latest_row(sheet_for_ticker(sheets, "Narrative", ticker), "generated_at")
    thesis = narrative_row.get("investment_thesis") if narrative_row else None
    if thesis and pd.notna(thesis):
        st.markdown("**Latest take:**")
        st.write(escape_markdown_math(thesis))
        st.caption("Full writeup in the AI Analysis tab.")

    last_updated = dash_row.get("last_updated")
    if last_updated is not None and pd.notna(last_updated):
        st.caption(f"Last updated: {fmt('last_updated', last_updated)}")


def render_table_tab(sheets: dict, tab_name: str, ticker: str, sort_col: str) -> None:
    df = sheet_for_ticker(sheets, tab_name, ticker)
    if df.empty:
        st.info(f"No {tab_name} data for {ticker} yet.")
        return
    df = df.sort_values(sort_col, ascending=False).drop(columns=["ticker"], errors="ignore")
    display = df.copy()
    for col in display.columns:
        display[col] = df[col].apply(lambda v, c=col: fmt(c, v))
    st.dataframe(display, width="stretch", hide_index=True, height=min(600, 60 + 35 * len(display)))


def render_news_tab(sheets: dict, ticker: str) -> None:
    df = sheet_for_ticker(sheets, "News", ticker)
    if df.empty:
        st.info(f"No news synced for {ticker} yet.")
        return
    df = df.sort_values("date", ascending=False)
    for _, item in df.iterrows():
        headline = escape_markdown_math(item.get("headline") or "(untitled)")
        link = item.get("link") or "#"
        summary = item.get("ai_summary")
        summary_html = (
            f'<div class="news-summary">{escape_markdown_math(summary)}</div>' if summary and pd.notna(summary) else ""
        )
        sentiment = item.get("sentiment")
        category = item.get("category")
        badges = ""
        if sentiment and pd.notna(sentiment):
            badges += badge(sentiment, SENTIMENT_COLORS.get(sentiment, CATEGORY_COLOR))
        if category and pd.notna(category):
            badges += badge(category, CATEGORY_COLOR)
        publisher = item.get("publisher") or ""
        date = fmt("date", item.get("date"))
        st.markdown(
            f"""<div class="news-card">
                <a href="{link}" target="_blank">{headline}</a>
                {summary_html}
                <div>{badges}<span class="news-meta">{publisher} &middot; {date}</span></div>
            </div>""",
            unsafe_allow_html=True,
        )


NARRATIVE_SECTIONS = [
    ("Executive Summary", "executive_summary"),
    ("Key Developments", "key_developments"),
    ("Bull Case", "bull_case"),
    ("Bear Case", "bear_case"),
    ("Financial Health", "financial_health"),
    ("Valuation Assessment", "valuation_assessment"),
    ("Competitive Position", "competitive_position"),
    ("Long-Term Outlook", "long_term_outlook"),
    ("Investment Thesis", "investment_thesis"),
]


def render_narrative_tab(sheets: dict, ticker: str) -> None:
    df = sheet_for_ticker(sheets, "Narrative", ticker)
    row = latest_row(df, "generated_at")
    if not row:
        st.info(f"No AI analysis generated for {ticker} yet.")
        return
    st.caption(f"Generated {fmt('generated_at', row.get('generated_at'))}")
    for title, field in NARRATIVE_SECTIONS:
        value = row.get(field)
        if value and pd.notna(value):
            with st.expander(title, expanded=(field in ("executive_summary", "investment_thesis"))):
                st.write(escape_markdown_math(value))
    confidence = row.get("confidence_score")
    if confidence is not None and pd.notna(confidence):
        st.divider()
        st.markdown(f"**Confidence: {fmt('confidence_score', confidence)} / 100**")
        rationale = row.get("confidence_rationale")
        if rationale and pd.notna(rationale):
            st.write(escape_markdown_math(rationale))


def _safe_float(value) -> float:
    try:
        if value is None or pd.isna(value):
            return 0.0
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def render_watchlist_tab(ticker: str) -> None:
    existing = excel_workbook.get_watchlist_record(ticker) or {}

    if not CAN_EDIT_REMOTELY:
        st.caption("Read-only here -- edit Watchlist notes from the app on your Mac.")
        for label, field in [
            ("Investment Thesis", "investment_thesis"), ("Catalysts", "catalysts"), ("Risks", "risks"),
            ("Target Price", "target_price"), ("Personal Rating", "personal_rating"), ("Notes", "notes"),
        ]:
            value = existing.get(field)
            if value and pd.notna(value):
                st.markdown(f"**{label}**")
                st.write(escape_markdown_math(value))
        return

    with st.form(f"watchlist_{ticker}"):
        thesis = st.text_area("Investment Thesis", value=existing.get("investment_thesis") or "")
        catalysts = st.text_area("Catalysts", value=existing.get("catalysts") or "")
        risks = st.text_area("Risks", value=existing.get("risks") or "")
        c1, c2 = st.columns(2)
        target_price = c1.number_input("Target Price ($)", value=_safe_float(existing.get("target_price")), step=0.5, format="%.2f")
        personal_rating = c2.text_input("Personal Rating", value=existing.get("personal_rating") or "")
        notes = st.text_area("Notes", value=existing.get("notes") or "")
        save = st.form_submit_button("Save")

    if save:
        if not IS_LOCAL_INSTANCE:
            git_sync.sync_before_write()
        excel_workbook.update_watchlist_manual_fields(
            ticker,
            investment_thesis=thesis or None,
            catalysts=catalysts or None,
            risks=risks or None,
            target_price=target_price or None,
            personal_rating=personal_rating or None,
            notes=notes or None,
        )
        git_sync.push_state_if_changed(f"Update Watchlist: {ticker}")
        st.cache_data.clear()
        st.success("Saved.")
        st.rerun()


def render_company_view(sheets: dict, ticker: str) -> None:
    tabs = st.tabs(["Overview", "Financials", "Ratios", "Valuation", "News", "AI Analysis", "Watchlist"])
    with tabs[0]:
        render_overview_tab(sheets, ticker)
    with tabs[1]:
        render_table_tab(sheets, "Financials", ticker, "period_end")
    with tabs[2]:
        render_table_tab(sheets, "Ratios", ticker, "period_end")
    with tabs[3]:
        render_table_tab(sheets, "Valuation", ticker, "as_of")
    with tabs[4]:
        render_news_tab(sheets, ticker)
    with tabs[5]:
        render_narrative_tab(sheets, ticker)
    with tabs[6]:
        render_watchlist_tab(ticker)


def main() -> None:
    st.session_state.setdefault("view_mode", "company")
    if "selected_ticker" not in st.session_state:
        active = tickers.get_all_tickers()
        st.session_state.selected_ticker = active[0] if active else None

    render_sidebar()
    sheets = load_workbook()

    if st.session_state.view_mode == "raw":
        render_raw_sheets(sheets)
        return

    ticker = st.session_state.get("selected_ticker")
    if not ticker:
        st.info("Add a company from the sidebar to get started.")
        return

    render_company_view(sheets, ticker)


if __name__ == "__main__":
    main()
