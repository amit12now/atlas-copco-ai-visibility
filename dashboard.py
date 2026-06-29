#!/usr/bin/env python3
"""
dashboard.py  (Atlas Copco build)
----------------------------------
Client-facing AI Visibility report for the AI Visibility & SEO Tracking template.

Reads ONLY the Clean_* sheets of AI_Visibility_Tracking_Template.xlsx (never the
RAW_* tabs) so it always reflects whatever the last run of update_template.py
produced. Run with:

    streamlit run dashboard.py

By default it looks for AI_Visibility_Tracking_Template.xlsx in the same folder
as this script. You can also point it at a different copy from the sidebar.

Layout: an "Overview" tab with the big plain-language numbers and a mix of
chart types (donut, gauge, styled bar charts), plus one simple tab per data
area for people who want to dig deeper. Every tab uses the same
plain-language, non-technical style -- no raw data tables, no jargon.

This client has no Sources export, so the "Sources" tab and the Overview's
"Sources AI relies on" section are omitted (the rest of the original
reusable template's layout/behavior is unchanged).

Visual theme uses the Atlas Copco brand palette (blue / gray / beige), sourced
from brandmanual.atlascopco.com. Colors and the brand display name are pulled
from Settings!Brand_Display_Name so this stays a reusable template for other
clients -- just change that one cell (and .streamlit/config.toml) to rebrand.
"""
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

DEFAULT_PATH = Path(__file__).parent / "AI_Visibility_Tracking_Template.xlsx"

CLEAN_SHEETS = [
    "Clean_AIO_Rankings", "Clean_Brand_Topics", "Clean_Gap_Topics",
    "Clean_Gap_Topics_Competitors", "Clean_Prompts", "Clean_Sources",
    "Clean_ChatGPT_Results", "Clean_GA4_Traffic", "Clean_GSC_Performance",
]

BLUE = "#0092BC"
BLUE_DARK = "#006F8F"
BLUE_TINT = "#F5FBFC"
BLUE_LIGHT = "#7AC6DC"
BEIGE = "#DFD6CE"
GREY = "#5B6770"
GREY_LIGHT = "#D1D4D7"
TEXT = "#2F363A"
GOOD = "#2E7D32"
BAD = "#C0392B"
FLAT = "#8C8C8C"

DONUT_PALETTE = [BLUE, BEIGE, BLUE_LIGHT, GREY, GREY_LIGHT, BLUE_DARK]

# Used to keep the "topics AI knows you for" spotlight on-brand: topics tagged
# in the underlying SEO data sometimes include unrelated noise (e.g. an
# unrelated brand or industry that happens to share a keyword). For Atlas
# Copco we prefer topics that are clearly about compressors / industrial air
# & gas equipment. Tune this list if reusing the template for a different
# industry.
RELEVANT_TOPIC_KEYWORDS = [
    "atlas copco", "compressor", "air dryer", "nitrogen", "generator",
    "vacuum", "pump", "drill", "hydraulic", "pneumatic", "rotary screw",
    "industrial tool", "chiller", "boiler",
]

LLM_DISPLAY = {
    "chatgpt": "ChatGPT",
    "gemini": "Gemini",
    "google_ai_mode": "Google AI Mode",
    "google_ai_overview": "Google AI Overview",
}

CATEGORY_DISPLAY = {
    "cites_target": "Cited as a source",
    "mentions_target": "Mentioned by name",
    "mentions_target;cites_target": "Both mentioned & cited",
}

# Renames for topic labels that come through from the underlying SEO data
# looking too technical/awkward for a client-facing report. Add to this dict
# instead of editing the raw data, so it survives the next update_template.py run.
TOPIC_DISPLAY = {}

st.set_page_config(page_title="AI Visibility Report", page_icon="\U0001F3ED", layout="wide")

st.markdown(
    f"""
    <style>
    .block-container {{ padding-top: 1.2rem; max-width: 1040px; }}
    .report-header {{
        background: linear-gradient(135deg, {BLUE} 0%, {BLUE_DARK} 100%);
        color: white; padding: 26px 32px; border-radius: 14px; margin-bottom: 18px;
    }}
    .report-header h1 {{ color: white; margin: 0; font-size: 1.7rem; font-weight: 700; }}
    .report-header p {{ color: #E2F4F9; margin: 4px 0 0 0; font-size: 0.95rem; }}
    .verdict-box {{
        background: {BLUE_TINT}; border-left: 5px solid {BLUE_DARK}; border-radius: 10px;
        padding: 16px 20px; margin: 0 0 26px 0; font-size: 1.05rem; line-height:1.5; color: {TEXT};
    }}
    .scorecard {{
        background: white; border-radius: 12px; padding: 16px 18px;
        box-shadow: 0 1px 6px rgba(0,0,0,0.08); border-left: 5px solid {BLUE};
        min-height: 132px; display:flex; flex-direction:column; justify-content:space-between;
    }}
    .scorecard-icon {{ font-size: 1.3rem; }}
    .scorecard-label {{
        color:#6b6b6b; font-size:0.72rem; font-weight:700; text-transform:uppercase;
        letter-spacing:.04em; margin-top:2px;
    }}
    .scorecard-value {{ font-size: 1.85rem; font-weight:700; color:{TEXT}; line-height:1.1; margin-top:6px; }}
    .scorecard-sub {{ font-size:0.8rem; font-weight:600; line-height:1.35; margin-top:6px; }}
    .sc-good {{ color:{GOOD}; }}
    .sc-bad {{ color:{BAD}; }}
    .sc-flat {{ color:{FLAT}; }}
    .context-note {{ color:#8a8a8a; font-size:0.82rem; margin: 16px 0 0 0; }}
    .section-title {{ font-size:1.2rem; font-weight:700; color:{TEXT}; margin: 30px 0 2px 0; }}
    .section-sub {{ color:#6b6b6b; font-size:0.9rem; margin-bottom:10px; }}
    .chart-card {{
        background:white; border-radius:12px; padding:10px 16px 4px 16px; margin-bottom:6px;
        box-shadow: 0 1px 6px rgba(0,0,0,0.06);
    }}
    .topic-card {{
        background:white; border-radius:12px; padding:14px 18px; margin-bottom:10px;
        box-shadow: 0 1px 6px rgba(0,0,0,0.06);
    }}
    .topic-name {{ font-size:0.98rem; font-weight:700; color:{TEXT}; margin-bottom:8px; }}
    .topic-track {{ background:{GREY_LIGHT}; border-radius:8px; height:10px; overflow:hidden; margin-bottom:6px; }}
    .topic-fill {{ height:100%; border-radius:8px; background:{BLUE}; }}
    .topic-note {{ color:#6b6b6b; font-size:0.82rem; }}
    .mention-row {{
        display:flex; align-items:center; gap:10px; padding:9px 0;
        border-bottom: 1px solid {GREY_LIGHT}; font-size:0.92rem;
    }}
    .mention-row:last-child {{ border-bottom:none; }}
    .mention-ok {{ color:{GOOD}; font-weight:700; }}
    .mention-no {{ color:{BAD}; font-weight:700; }}
    .empty-note {{ color:#9a9a9a; font-size:0.88rem; font-style:italic; margin-bottom:20px; }}
    .footer-note {{
        color:#8a8a8a; font-size:0.82rem; margin-top:40px; padding-top:14px;
        border-top:1px solid {GREY_LIGHT};
    }}
    .stTabs [data-baseweb="tab"] {{ font-weight:600; }}
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner="Loading template...")
def load_data(file_bytes, mtime):
    sheets = {}
    for name in CLEAN_SHEETS:
        try:
            sheets[name] = pd.read_excel(file_bytes, sheet_name=name)
        except ValueError:
            sheets[name] = pd.DataFrame()
    try:
        settings = pd.read_excel(file_bytes, sheet_name="Settings")
        settings = dict(zip(settings["Setting"], settings["Value"]))
    except Exception:
        settings = {}
    return sheets, settings


def latest_slice(df, month_col="Report_Month"):
    """Return only the rows for the most recent month, plus the month value itself."""
    if df.empty or month_col not in df.columns:
        return df, None
    m = df[month_col].dropna().max()
    if pd.isna(m):
        return df, None
    return df[df[month_col] == m], m


def latest_and_prev(df, month_col):
    if df.empty or month_col not in df.columns:
        return None, None
    months = sorted(df[month_col].dropna().unique())
    if not months:
        return None, None
    latest = months[-1]
    prev = months[-2] if len(months) > 1 else None
    return latest, prev


def ordinal(n):
    n = int(n)
    if 10 <= n % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def month_label(ym):
    try:
        return pd.Period(str(ym), freq="M").strftime("%B %Y")
    except Exception:
        return str(ym)


def short_url(u):
    """Trim a full URL down to just its path, for compact display."""
    try:
        path = urlparse(str(u)).path
        return path if path and path != "/" else str(u)
    except Exception:
        return str(u)


def llm_name(code):
    return LLM_DISPLAY.get(str(code).strip().lower(), str(code).replace("_", " ").title())


def category_name(code):
    return CATEGORY_DISPLAY.get(str(code).strip(), str(code).replace("_", " ").title())


def topic_name(topic):
    return TOPIC_DISPLAY.get(str(topic).strip(), str(topic))


def is_relevant_topic(text):
    t = str(text).lower()
    return any(k in t for k in RELEVANT_TOPIC_KEYWORDS)


def trend_text(delta, higher_is_better=True, unit="pts", compare="last month"):
    """Plain-language up/down/same sentence + good/bad/flat class for a numeric delta."""
    if delta is None or (isinstance(delta, float) and pd.isna(delta)):
        return "First month tracked", "flat"
    if round(delta, 1) == 0:
        return f"Same as {compare}", "flat"
    cls = "good" if (delta > 0) == higher_is_better else "bad"
    word = "Up" if delta > 0 else "Down"
    return f"{word} {abs(delta):.0f}{unit} vs {compare}", cls


def rank_trend(rank_now, rank_prev):
    if rank_prev is None or rank_now is None:
        return "Among tracked competitors", "flat"
    if rank_now < rank_prev:
        return f"Up from {ordinal(rank_prev)} last month", "good"
    if rank_now > rank_prev:
        return f"Down from {ordinal(rank_prev)} last month", "bad"
    return "Same as last month", "flat"


def scorecard(icon, label, value, sub_text, sub_cls):
    st.markdown(
        f"""
        <div class="scorecard">
            <div>
                <div class="scorecard-icon">{icon}</div>
                <div class="scorecard-label">{label}</div>
            </div>
            <div>
                <div class="scorecard-value">{value}</div>
                <div class="scorecard-sub sc-{sub_cls}">{sub_text}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_title(title, subtitle):
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="section-sub">{subtitle}</div>', unsafe_allow_html=True)


def bar_chart(labels, values, highlight_index=0, value_text=None, height=None):
    """Styled horizontal bar chart -- used for ranked / leaderboard-style data."""
    n = len(labels)
    colors = [BLUE if i == highlight_index else GREY for i in range(n)]
    text = value_text if value_text is not None else [str(v) for v in values]
    fig = go.Figure(
        go.Bar(
            x=values, y=labels, orientation="h",
            marker=dict(color=colors),
            text=text, textposition="outside", cliponaxis=False,
            textfont=dict(size=12.5, color=TEXT),
        )
    )
    fig.update_layout(
        height=height or max(46 * n + 30, 110),
        margin=dict(l=10, r=40, t=8, b=8),
        xaxis=dict(visible=False, range=[0, max(values) * 1.18 if values else 1]),
        yaxis=dict(autorange="reversed", tickfont=dict(size=12.5, color=TEXT)),
        plot_bgcolor="white", paper_bgcolor="white",
        showlegend=False,
        bargap=0.32,
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def donut_chart(labels, values, height=290, hole=0.55):
    """Pie/donut chart -- used for share-of-total data (platform mix, device mix, etc)."""
    colors = (DONUT_PALETTE * (len(labels) // len(DONUT_PALETTE) + 1))[: len(labels)]
    fig = go.Figure(
        go.Pie(
            labels=labels, values=values, hole=hole,
            marker=dict(colors=colors, line=dict(color="white", width=2)),
            textinfo="percent", textfont=dict(size=12.5),
            sort=False,
        )
    )
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=10, b=10),
        legend=dict(orientation="v", font=dict(size=12.5)),
        paper_bgcolor="white",
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def gauge_chart(value, max_value=100, suffix="%", height=180, label=None):
    """Radial gauge -- used to highlight one standout percentage."""
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=value,
            number={"suffix": suffix, "font": {"size": 30, "color": TEXT}},
            gauge={
                "axis": {"range": [0, max_value], "visible": False},
                "bar": {"color": BLUE, "thickness": 0.32},
                "bgcolor": GREY_LIGHT,
                "borderwidth": 0,
            },
        )
    )
    fig.update_layout(height=height, margin=dict(l=20, r=20, t=10, b=0), paper_bgcolor="white")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    if label:
        st.markdown(f"<div style='text-align:center; color:#6b6b6b; font-size:0.85rem; margin-top:-14px;'>{label}</div>", unsafe_allow_html=True)


def vbar_chart(labels, values, highlight_index=-1, value_text=None, height=260):
    """Vertical bar chart -- used for month-by-month trend totals (e.g. visits over time)."""
    n = len(labels)
    colors = [BLUE if i == (highlight_index if highlight_index >= 0 else n - 1) else GREY for i in range(n)]
    text = value_text if value_text is not None else [str(v) for v in values]
    fig = go.Figure(
        go.Bar(
            x=labels, y=values,
            marker=dict(color=colors),
            text=text, textposition="outside", cliponaxis=False,
            textfont=dict(size=12.5, color=TEXT),
        )
    )
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=28, b=8),
        yaxis=dict(visible=False, range=[0, max(values) * 1.25 if values else 1]),
        xaxis=dict(tickfont=dict(size=12.5, color=TEXT)),
        plot_bgcolor="white", paper_bgcolor="white",
        showlegend=False,
        bargap=0.35,
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def line_chart(x, series, height=320):
    """series: dict of {name: (values, color)}. Used for trends over time."""
    fig = go.Figure()
    for name, (values, color) in series.items():
        fig.add_trace(go.Scatter(
            x=x, y=values, name=name, mode="lines+markers",
            line=dict(color=color, width=3), marker=dict(size=7),
        ))
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, font=dict(size=12.5)),
        xaxis=dict(showgrid=False, tickfont=dict(size=12)),
        yaxis=dict(showgrid=True, gridcolor=GREY_LIGHT, tickfont=dict(size=12)),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def topic_card(name, pct_value, note):
    st.markdown(
        f"""
        <div class="topic-card">
            <div class="topic-name">{name}</div>
            <div class="topic-track"><div class="topic-fill" style="width:{max(pct_value, 2):.0f}%;"></div></div>
            <div class="topic-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def mention_list(rows):
    """rows: list of (text, mentioned_bool)."""
    html = []
    for text, mentioned in rows:
        if mentioned:
            html.append(f'<div class="mention-row"><span class="mention-ok">✓ Mentioned</span><span>&mdash; "{text}"</span></div>')
        else:
            html.append(f'<div class="mention-row"><span class="mention-no">✗ Not mentioned</span><span>&mdash; "{text}"</span></div>')
    st.markdown("".join(html), unsafe_allow_html=True)


def empty_note(text):
    st.markdown(f'<div class="empty-note">{text}</div>', unsafe_allow_html=True)


# ------------------------------------------------------------------- Sidebar
st.sidebar.markdown(
    f"<h2 style='color:{BLUE}; margin-bottom:0;'>\U0001F3ED AI Visibility</h2>"
    f"<p style='color:#777; margin-top:0;'>Client report</p>",
    unsafe_allow_html=True,
)
upload = st.sidebar.file_uploader("Or upload a template copy", type=["xlsx"])
source = upload if upload is not None else (str(DEFAULT_PATH) if DEFAULT_PATH.exists() else None)

if source is None:
    st.warning(
        f"Couldn't find {DEFAULT_PATH.name} next to dashboard.py. "
        "Upload a copy from the sidebar, or place the template in this folder."
    )
    st.stop()

mtime = upload.size if upload is not None else DEFAULT_PATH.stat().st_mtime
sheets, settings = load_data(source, mtime)

brand_domain = str(settings.get("Brand_Domain", "")).strip().lower()
brand_label = str(settings.get("Brand_Display_Name") or brand_domain or "Brand").strip()

st.sidebar.markdown(
    f"<div style='background:{BLUE_TINT}; color:{BLUE_DARK}; border-radius:8px; padding:6px 12px; margin-bottom:6px; font-size:0.85rem;'>"
    f"\U0001F3E2 {brand_label}</div>"
    f"<div style='background:{BLUE_TINT}; color:{BLUE_DARK}; border-radius:8px; padding:6px 12px; font-size:0.85rem;'>"
    f"\U0001F551 Last updated: {settings.get('Last_Updated', 'n/a')}</div>",
    unsafe_allow_html=True,
)

aio = sheets["Clean_AIO_Rankings"]
brand_topics = sheets["Clean_Brand_Topics"]
prompts = sheets["Clean_Prompts"]
sources = sheets["Clean_Sources"]
chatgpt = sheets["Clean_ChatGPT_Results"]
ga4 = sheets["Clean_GA4_Traffic"]
gsc = sheets["Clean_GSC_Performance"]

# ---------------------------------------------------------------- Header
report_month_setting = settings.get("Report_Month")
header_month = month_label(report_month_setting) if report_month_setting else ""
st.markdown(
    f"""
    <div class="report-header">
        <h1>\U0001F3ED AI Visibility Report</h1>
        <p>{brand_label}{' &middot; ' + header_month if header_month else ''}</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------- Core numbers
aio_latest, aio_prev = latest_and_prev(aio, "Report_Month")
brand_now = aio[(aio["Report_Month"] == aio_latest) & (aio["Is_Brand"])] if aio_latest else pd.DataFrame()
top3_now = int((brand_now["Position"] <= 3).sum()) if not brand_now.empty else 0
total_now = len(brand_now)
pct_now = (top3_now / total_now * 100) if total_now else None

d_pct = None
if aio_prev is not None:
    brand_prev = aio[(aio["Report_Month"] == aio_prev) & (aio["Is_Brand"])]
    top3_prev, total_prev = int((brand_prev["Position"] <= 3).sum()), len(brand_prev)
    pct_prev = (top3_prev / total_prev * 100) if total_prev else None
    if pct_prev is not None and pct_now is not None:
        d_pct = pct_now - pct_prev

domain_avg_now = (
    aio[aio["Report_Month"] == aio_latest].dropna(subset=["Position"]).groupby("Domain")["Position"].mean().sort_values()
    if aio_latest else pd.Series(dtype=float)
)
total_domains = len(domain_avg_now)
rank_now = (list(domain_avg_now.index).index(brand_domain) + 1) if brand_domain in domain_avg_now.index else None

rank_prev = None
if aio_prev is not None:
    domain_avg_prev = aio[aio["Report_Month"] == aio_prev].dropna(subset=["Position"]).groupby("Domain")["Position"].mean().sort_values()
    if brand_domain in domain_avg_prev.index:
        rank_prev = list(domain_avg_prev.index).index(brand_domain) + 1

chat_latest, chat_prev = latest_and_prev(chatgpt, "Report_Month")
chat_now_rows = chatgpt[chatgpt["Report_Month"] == chat_latest] if chat_latest else pd.DataFrame()
rate_now = (chat_now_rows["Brand_Mentioned"].mean() * 100) if not chat_now_rows.empty else None
n_prompts = len(chat_now_rows)
d_rate = None
if chat_prev is not None and rate_now is not None:
    rate_prev = chatgpt.loc[chatgpt["Report_Month"] == chat_prev, "Brand_Mentioned"].mean() * 100
    d_rate = rate_now - rate_prev

ga_latest, ga_prev = latest_and_prev(ga4, "Year_Month")
sess_now = ga4.loc[ga4["Year_Month"] == ga_latest, "Sessions"].sum() if ga_latest else 0
is_to_date = bool(ga4.loc[ga4["Year_Month"] == ga_latest, "Is_To_Date"].any()) if ga_latest and not ga4.empty else False
d_sess_pct = None
if ga_prev is not None:
    sess_prev = ga4.loc[ga4["Year_Month"] == ga_prev, "Sessions"].sum()
    if sess_prev:
        d_sess_pct = (sess_now - sess_prev) / sess_prev * 100

# =====================================================================
# TABS  (no "Sources" tab for this client -- no Sources export available)
# =====================================================================
tab_overview, tab_rankings, tab_topics, tab_chatgpt, tab_traffic, tab_gsc = st.tabs(
    ["\U0001F3E0 Overview", "\U0001F50E AI Search Rankings", "\U0001F9E0 Topics & Prompts",
     "\U0001F4AC ChatGPT Spot-Checks", "\U0001F4C8 Website Traffic",
     "\U0001F30D Google AI Search"]
)

# --------------------------------------------------------------- OVERVIEW
with tab_overview:
    verdict_parts = []
    if rank_now and total_domains > 1:
        if rank_now == 1:
            verdict_parts.append(f"This month, <b>{brand_label}</b> is the most visible name among the {total_domains} competitors tracked in AI search.")
        else:
            verdict_parts.append(f"This month, <b>{brand_label}</b> ranks {ordinal(rank_now)} out of {total_domains} tracked names in AI search.")
    elif aio_latest:
        verdict_parts.append(f"This month, <b>{brand_label}</b> appeared in the AI Overview for {top3_now} of {total_now} tracked keywords.")
    if chat_latest and rate_now is not None:
        verdict_parts.append(f"Asked directly, ChatGPT mentioned {brand_label} in {rate_now:.0f}% of the questions we tested.")
    if ga_latest and d_sess_pct is not None and d_sess_pct > 0:
        verdict_parts.append(f"AI-driven website visits are up {d_sess_pct:.0f}% vs last month.")
    if verdict_parts:
        st.markdown(f'<div class="verdict-box">\U0001F44B {" ".join(verdict_parts)}</div>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        sub, cls = rank_trend(rank_now, rank_prev)
        scorecard("\U0001F3C6", "Rank in AI search", f"{ordinal(rank_now)} of {total_domains}" if rank_now else "n/a", sub, cls)
    with c2:
        sub, cls = trend_text(d_pct, higher_is_better=True, unit=" pts")
        scorecard("\U0001F50E", "Shows up in AI search", f"{pct_now:.0f}%" if pct_now is not None else "n/a", sub, cls)
    with c3:
        sub, cls = trend_text(d_rate, higher_is_better=True, unit=" pts")
        scorecard("\U0001F4AC", "Mentioned by ChatGPT", f"{rate_now:.0f}%" if rate_now is not None else "n/a", sub, cls)
    with c4:
        sub, cls = trend_text(d_sess_pct, higher_is_better=True, unit="%")
        if is_to_date:
            sub += " (estimate, month in progress)"
        scorecard("\U0001F916", "Visits from AI chatbots", f"{int(sess_now):,}" if ga_latest else "n/a", sub, cls)

    if chat_latest and n_prompts:
        st.markdown(
            f'<div class="context-note">Based on {n_prompts} real questions tested across ChatGPT, Gemini and other AI assistants this month.</div>',
            unsafe_allow_html=True,
        )

    section_title("Website visits from AI chatbots", "How many people clicked through to your website from an AI chat answer, month by month.")
    if ga4.empty:
        empty_note("No AI-referral traffic data yet.")
    else:
        months = sorted(ga4["Year_Month"].unique())
        month_labels = [month_label(m) for m in months]
        totals = [int(ga4.loc[ga4["Year_Month"] == m, "Sessions"].sum()) for m in months]
        vbar_chart(month_labels, totals, value_text=[f"{v:,}" for v in totals])
        if is_to_date:
            empty_note("This month's figure is an estimate -- the month is still in progress.")

    section_title("How you compare in AI search", "When someone asks an AI assistant a question in your category, this shows who it tends to mention first.")
    if domain_avg_now.empty:
        empty_note("No AI search ranking data yet.")
    else:
        ranked = list(domain_avg_now.items())
        labels = [brand_label if d == brand_domain else d for d, _ in ranked]
        scores = [round(1 / p, 3) for _, p in ranked]
        texts = [f"{ordinal(i + 1)} place" for i in range(len(ranked))]
        hi = next((i for i, (d, _) in enumerate(ranked) if d == brand_domain), 0)
        bar_chart(labels, scores, highlight_index=hi, value_text=texts)

    col_a, col_b = st.columns([3, 2])
    with col_a:
        section_title("What AI already knows you for", "The topics AI most often connects to your name when answering related questions.")
        topics_now, topics_month = latest_slice(brand_topics, "Report_Month")
        if topics_now.empty:
            empty_note("No brand-topic data yet.")
        else:
            relevant = topics_now[topics_now["Topic"].apply(is_relevant_topic)]
            pool = relevant if len(relevant) >= 3 else topics_now
            top_topics = pool.sort_values("Visibility", ascending=False).head(3)
            max_vis = top_topics["Visibility"].max()
            for i, row in enumerate(top_topics.itertuples()):
                note = "Strongest topic this month." if i == 0 else "Also well-recognized by AI."
                topic_card(f'"{topic_name(row.Topic)}"', row.Visibility / max_vis * 100, note)
    with col_b:
        if not topics_now.empty:
            st.markdown("<div style='height:38px'></div>", unsafe_allow_html=True)
            top_row = top_topics.iloc[0]
            gauge_chart(round(float(top_row["Visibility"]), 0), max_value=100, suffix="", height=190, label=f'AI-visibility score for "{topic_name(top_row["Topic"])}"')

    col_c, col_d = st.columns(2)
    with col_c:
        section_title("What we asked AI assistants directly", "Share of the real prompts we tested across each AI assistant this month.")
        prompts_now, _ = latest_slice(prompts, "Report_Month")
        if prompts_now.empty:
            empty_note("No prompt-testing data yet.")
        else:
            counts = prompts_now["LLM"].value_counts()
            donut_chart([llm_name(x) for x in counts.index], counts.values.tolist(), height=270)
    with col_d:
        section_title("Where AI-driven visits come from", "Share of AI-chatbot website visits this month, by platform.")
        ga_now, _ = latest_slice(ga4, "Year_Month")
        if ga_now.empty:
            empty_note("No AI-referral traffic data yet.")
        else:
            by_platform = ga_now.groupby("Platform")["Sessions"].sum().sort_values(ascending=False)
            donut_chart(list(by_platform.index), by_platform.values.tolist(), height=270)
            if is_to_date:
                empty_note("This month's figures are an estimate -- the month is still in progress.")

    if not chat_now_rows.empty:
        section_title("ChatGPT spot-checks", f"A sample of the real questions we typed into ChatGPT this month, and whether it mentioned {brand_label} by name.")
        mentioned = chat_now_rows[chat_now_rows["Brand_Mentioned"]].head(3)
        not_mentioned = chat_now_rows[~chat_now_rows["Brand_Mentioned"]].head(2)
        sample = pd.concat([mentioned, not_mentioned])
        mention_list([(row.Keyword, bool(row.Brand_Mentioned)) for row in sample.itertuples()])

    st.markdown(
        '<div class="footer-note">This is the at-a-glance summary. Use the tabs above to dig into each area in more detail.</div>',
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------- AI SEARCH RANKINGS
with tab_rankings:
    section_title("How AI ranks you vs. competitors", f"Average position across {total_now} tracked keywords this month, lower is better.")
    if domain_avg_now.empty:
        empty_note("No AI search ranking data yet.")
    else:
        ranked = list(domain_avg_now.items())
        labels = [brand_label if d == brand_domain else d for d, _ in ranked]
        positions = [round(p, 1) for _, p in ranked]
        texts = [f"avg. position {p:.1f}" for p in positions]
        hi = next((i for i, (d, _) in enumerate(ranked) if d == brand_domain), 0)
        bar_chart(labels, [round(1 / p, 3) for p in positions], highlight_index=hi, value_text=texts)

    col1, col2 = st.columns([3, 2])
    with col1:
        section_title("Keywords where you rank #1", "The search questions where AI is most likely to put your name first.")
        win = brand_now.dropna(subset=["Position"])
        win = win[win["Position"] == 1].sort_values("Search_Volume", ascending=False).head(8) if not win.empty else win
        if win.empty:
            empty_note("No #1 keyword rankings yet.")
        else:
            bar_chart(win["Keyword"].tolist(), [1.0] * len(win), highlight_index=0, value_text=["#1"] * len(win), height=38 * len(win) + 30)

        section_title("Keywords with room to grow", "Tracked keywords where AI ranks you lower -- the best opportunities to improve.")
        grow = brand_now.dropna(subset=["Position"])
        grow = grow[grow["Position"] > 3].sort_values("Position", ascending=False).head(5) if not grow.empty else grow
        if grow.empty:
            empty_note("No lower-ranked keywords to highlight -- great coverage!")
        else:
            bar_chart(
                grow["Keyword"].tolist(),
                [round(1 / p, 3) for p in grow["Position"]],
                highlight_index=None if True else 0,
                value_text=[f"#{int(p)}" for p in grow["Position"]],
                height=38 * len(grow) + 30,
            )
    with col2:
        section_title("Your position mix", "How your tracked keywords break down by AI search position.")
        if brand_now.empty:
            empty_note("No keyword-level ranking data yet.")
        else:
            pos = brand_now["Position"]
            b1 = int((pos == 1).sum())
            b2 = int(((pos >= 2) & (pos <= 3)).sum())
            b3 = int(((pos >= 4) & (pos <= 10)).sum())
            b4 = int(pos.isna().sum() + (pos > 10).sum())
            labels, values = [], []
            for lab, val in [("Position 1", b1), ("Position 2-3", b2), ("Position 4-10", b3), ("Not in top 10", b4)]:
                if val > 0:
                    labels.append(lab)
                    values.append(val)
            donut_chart(labels, values, height=300)

# ---------------------------------------------------------- TOPICS & PROMPTS
with tab_topics:
    section_title("Topics AI connects you with", "The themes AI most often associates with your name, ranked by visibility.")
    topics_now, _ = latest_slice(brand_topics, "Report_Month")
    if topics_now.empty:
        empty_note("No brand-topic data yet.")
    else:
        relevant = topics_now[topics_now["Topic"].apply(is_relevant_topic)]
        pool = relevant if len(relevant) >= 5 else topics_now
        top8 = pool.sort_values("Visibility", ascending=False).head(8)
        bar_chart([topic_name(t) for t in top8["Topic"]], top8["Visibility"].tolist(), highlight_index=0, height=40 * len(top8) + 30)

    col1, col2 = st.columns(2)
    with col1:
        section_title("How AI assistants refer to you", "When you come up in an AI answer, is it a name-drop, a citation, or both?")
        prompts_now, _ = latest_slice(prompts, "Report_Month")
        if prompts_now.empty:
            empty_note("No prompt-testing data yet.")
        else:
            cats = prompts_now["Category"].value_counts()
            donut_chart([category_name(x) for x in cats.index], cats.values.tolist())
    with col2:
        section_title("Prompts tested per AI assistant", "Real prompts we tested this month, covering the kinds of questions a buyer or specifier might ask.")
        if prompts_now.empty:
            empty_note("No prompt-testing data yet.")
        else:
            counts = prompts_now["LLM"].value_counts()
            donut_chart([llm_name(x) for x in counts.index], counts.values.tolist())

# ---------------------------------------------------------- CHATGPT SPOT-CHECKS
with tab_chatgpt:
    section_title("How often ChatGPT mentions you", f"Out of {n_prompts} real questions we asked ChatGPT this month.")
    if chat_now_rows.empty:
        empty_note("No ChatGPT spot-check data yet.")
    else:
        mentioned_n = int(chat_now_rows["Brand_Mentioned"].sum())
        donut_chart(["Mentioned", "Not mentioned"], [mentioned_n, n_prompts - mentioned_n], height=260)

        section_title("Every question we asked", f"All {n_prompts} questions we typed into ChatGPT this month, and whether it mentioned {brand_label} by name.")
        ordered = pd.concat([chat_now_rows[chat_now_rows["Brand_Mentioned"]], chat_now_rows[~chat_now_rows["Brand_Mentioned"]]])
        mention_list([(row.Keyword, bool(row.Brand_Mentioned)) for row in ordered.itertuples()])

# ---------------------------------------------------------- WEBSITE TRAFFIC
with tab_traffic:
    section_title("AI-driven website visits over time", "Sessions that arrived at your website straight from an AI chat answer, by platform, each month.")
    if ga4.empty:
        empty_note("No AI-referral traffic data yet.")
    else:
        months = sorted(ga4["Year_Month"].unique())
        month_labels = [month_label(m) for m in months]
        platforms = ga4["Platform"].unique().tolist()
        platform_colors = {p: DONUT_PALETTE[i % len(DONUT_PALETTE)] for i, p in enumerate(platforms)}
        series = {}
        for p in platforms:
            vals = [ga4.loc[(ga4["Year_Month"] == m) & (ga4["Platform"] == p), "Sessions"].sum() for m in months]
            series[p] = (vals, platform_colors[p])
        line_chart(month_labels, series)

        total_sess = sum(s[0][-1] for s in series.values())
        st.markdown(
            f'<div class="context-note">{int(total_sess):,} AI-driven visits in {month_labels[-1]}'
            f'{" (estimate, month in progress)" if is_to_date else ""}.</div>',
            unsafe_allow_html=True,
        )

    section_title("This month's platform mix", f"Where this month's {int(sess_now):,} AI-driven visits came from.")
    ga_now, _ = latest_slice(ga4, "Year_Month")
    if ga_now.empty:
        empty_note("No AI-referral traffic data yet.")
    else:
        by_platform = ga_now.groupby("Platform")["Sessions"].sum().sort_values(ascending=False)
        donut_chart(list(by_platform.index), by_platform.values.tolist())

# ---------------------------------------------------- GOOGLE AI OVERVIEWS (GSC)
with tab_gsc:
    section_title("Google AI Overviews & AI Mode performance", "How often your pages appeared in Google's AI Overviews and AI Mode results (Search Console's AI search reporting) over the last 28 days.")
    gsc_now, _ = latest_slice(gsc, "Report_Month")
    if gsc_now.empty:
        empty_note("No Search Console data yet.")
    else:
        pages = gsc_now[gsc_now["Dimension_Type"] == "Page"].sort_values("Impressions_Last_28d", ascending=False).head(10)
        if pages.empty:
            empty_note("No page-level Search Console data yet.")
        else:
            labels = [short_url(u) for u in pages["Dimension_Value"]]
            texts = [f"{int(v):,} {'▲' if c > 0 else ('▼' if c < 0 else '■')}" for v, c in zip(pages["Impressions_Last_28d"], pages["Change_Pct"])]
            bar_chart(labels, pages["Impressions_Last_28d"].tolist(), highlight_index=0, value_text=texts, height=38 * len(pages) + 30)

        col1, col2 = st.columns(2)
        with col1:
            section_title("Where in the world people find you", "Countries generating the most AI Overview / AI Mode impressions in the last 28 days.")
            countries = gsc_now[gsc_now["Dimension_Type"] == "Country"].sort_values("Impressions_Last_28d", ascending=False)
            if countries.empty:
                empty_note("No country-level Search Console data yet.")
            else:
                top5 = countries.head(5)
                other_sum = countries["Impressions_Last_28d"].iloc[5:].sum()
                labels = top5["Dimension_Value"].tolist()
                values = top5["Impressions_Last_28d"].tolist()
                if other_sum > 0:
                    labels.append("Other countries")
                    values.append(int(other_sum))
                donut_chart(labels, values)
        with col2:
            section_title("What device people search from", "AI Overview / AI Mode impressions in the last 28 days, by device type.")
            devices = gsc_now[gsc_now["Dimension_Type"] == "Device"].sort_values("Impressions_Last_28d", ascending=False)
            if devices.empty:
                empty_note("No device-level Search Console data yet.")
            else:
                donut_chart(devices["Dimension_Value"].tolist(), devices["Impressions_Last_28d"].tolist())
