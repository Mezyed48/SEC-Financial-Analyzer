"""
Streamlit Web Application: SEC 10-K Financial & Sentiment Analyzer

This local web frontend wraps the verified Steps 2-9 financial analysis pipeline,
providing an executive-style financial dashboard for:
  - Form 10-K retrieval from SEC EDGAR
  - Item 7 MD&A plain-text extraction and cleaning
  - Financial statements extraction (Income Statement, Balance Sheet, Cash Flows)
  - Loughran-McDonald financial sentiment analysis (8 sentiment categories, P/N ratio)
  - Historical split-adjusted stock prices (Yahoo Finance via yfinance)
  - 10-Year financial ratio analysis (Liquidity, Solvency, Efficiency, Profitability)
  - Long-term revenue & earnings growth trends
  - Apple vs. Tesla side-by-side executive comparison dashboard
  - One-click master Excel workbook (.xlsx) download

Runs on:
  streamlit run streamlit_app.py
"""

import os
from pathlib import Path
import re
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# Import the verified backend pipeline and dictionary functions from main.py
import main
from main import (
    run_analysis_workflow,
    PROJECT_ROOT,
    setup_dictionaries_directory,
    find_dictionary_file,
    load_loughran_mcdonald_dictionary
)


# ==============================================================================
# PAGE CONFIGURATION
# ==============================================================================

st.set_page_config(
    page_title="SEC 10-K Financial & Sentiment Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ==============================================================================
# SESSION STATE & THEME CONFIGURATION
# ==============================================================================

if "theme_mode" not in st.session_state:
    st.session_state.theme_mode = "🌙 Dark"

if "analysis_cache" not in st.session_state:
    st.session_state.analysis_cache = {}

if "company_mode" not in st.session_state:
    st.session_state.company_mode = "Apple (AAPL)"

if "custom_ticker" not in st.session_state:
    st.session_state.custom_ticker = "AAPL"

if "active_ticker" not in st.session_state:
    st.session_state.active_ticker = "AAPL"

if "sidebar_ticker_input" not in st.session_state:
    st.session_state.sidebar_ticker_input = "AAPL"

if "selected_filing_count" not in st.session_state:
    st.session_state.selected_filing_count = 4

is_dark = (st.session_state.theme_mode == "🌙 Dark")

if is_dark:
    THEME = {
        "is_dark": True,
        "name": "dark",
        "bg_color": "#0E1117",
        "sidebar_bg": "#171B24",
        "card_bg": "#161B22",
        "card_secondary_bg": "#202632",
        "header_card_bg": "#1C212B",
        "text_primary": "#F5F7FA",
        "text_secondary": "#AAB2C0",
        "text_muted": "#718096",
        "accent_blue": "#4C8DFF",
        "accent_blue_hover": "#6BA4FF",
        "border_color": "#2A3342",
        "positive": "#27C499",
        "negative": "#FF5C5C",
        "warning": "#F2C94C",
        "badge_aapl_bg": "rgba(76, 141, 255, 0.15)",
        "badge_aapl_text": "#7CB0FF",
        "badge_aapl_border": "#2A3342",
        "badge_sec_bg": "#202632",
        "badge_sec_text": "#AAB2C0",
        "badge_sec_border": "#2A3342",
        "intro_bg": "#161B22",
        "intro_border": "#4C8DFF",
        "intro_text": "#E2E8F0",
        "metric_cat_color": "#AAB2C0",
        "metric_name_color": "#F5F7FA",
        "metric_val_color": "#4C8DFF",
        "aapl_color": "#4C8DFF",
        "tsla_color": "#FF7A59",
        "plotly_template": "plotly_dark",
        "chart_paper_bg": "#0E1117",
        "chart_plot_bg": "#161B22",
        "chart_font_color": "#F5F7FA",
        "chart_grid_color": "#2A3342",
        "chart_zeroline_color": "#3B4758",
        "chart_neutral_line": "#718096",
        "chart_primary_line": "#4C8DFF",
        "chart_secondary_line": "#FF7A59",
        "chart_positive_color": "#27C499",
        "chart_negative_color": "#FF5C5C",
    }
else:
    THEME = {
        "is_dark": False,
        "name": "light",
        "bg_color": "#FFFFFF",
        "sidebar_bg": "#F8FAFC",
        "card_bg": "#FFFFFF",
        "card_secondary_bg": "#EDF2F7",
        "header_card_bg": "#F8FAFC",
        "text_primary": "#2D3748",
        "text_secondary": "#4A5568",
        "text_muted": "#718096",
        "accent_blue": "#1F497D",
        "accent_blue_hover": "#2B6CB0",
        "border_color": "#E2E8F0",
        "positive": "#2F855A",
        "negative": "#C53030",
        "warning": "#D69E2E",
        "badge_aapl_bg": "#EBF8FF",
        "badge_aapl_text": "#2B6CB0",
        "badge_aapl_border": "#BEE3F8",
        "badge_sec_bg": "#EDF2F7",
        "badge_sec_text": "#4A5568",
        "badge_sec_border": "#E2E8F0",
        "intro_bg": "#F7FAFC",
        "intro_border": "#1F497D",
        "intro_text": "#2D3748",
        "metric_cat_color": "#718096",
        "metric_name_color": "#2D3748",
        "metric_val_color": "#1F497D",
        "aapl_color": "#1F497D",
        "tsla_color": "#E4572E",
        "plotly_template": "plotly_white",
        "chart_paper_bg": "#FFFFFF",
        "chart_plot_bg": "#FFFFFF",
        "chart_font_color": "#2D3748",
        "chart_grid_color": "#EDF2F7",
        "chart_zeroline_color": "#CBD5E0",
        "chart_neutral_line": "#A0AEC0",
        "chart_primary_line": "#1F497D",
        "chart_secondary_line": "#E4572E",
        "chart_positive_color": "#2B6CB0",
        "chart_negative_color": "#C53030",
    }


# ==============================================================================
# DYNAMIC THEME STYLES (DARK / LIGHT COMPATIBLE)
# ==============================================================================

st.markdown(f"""
<style>
    /* Global Root Variables & Theme Colors */
    :root {{
        --background-color: {THEME['bg_color']};
        --secondary-background-color: {THEME['card_bg']};
        --text-color: {THEME['text_primary']};
        --primary-color: {THEME['accent_blue']};
    }}

    /* Main Application Body */
    .stApp {{
        background-color: {THEME['bg_color']};
        color: {THEME['text_primary']};
    }}
    [data-testid="stAppViewContainer"] {{
        background-color: {THEME['bg_color']};
    }}
    [data-testid="stHeader"] {{
        background-color: {THEME['bg_color']};
    }}

    /* Sidebar Styling */
    section[data-testid="stSidebar"] {{
        background-color: {THEME['sidebar_bg']} !important;
        border-right: 1px solid {THEME['border_color']} !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {{
        color: {THEME['text_primary']};
    }}
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3 {{
        color: {THEME['accent_blue']} !important;
    }}

    /* Titles and Subtitles */
    .main-title {{
        font-size: 2.1rem;
        font-weight: 700;
        color: {THEME['accent_blue']};
        margin-bottom: 0.2rem;
    }}
    .sub-title {{
        font-size: 1.05rem;
        font-weight: 400;
        color: {THEME['text_secondary']};
        margin-bottom: 0.8rem;
    }}
    h1, h2, h3, h4, h5, h6 {{
        color: {THEME['accent_blue']} !important;
    }}

    /* Intro Box */
    .intro-box {{
        background-color: {THEME['intro_bg']};
        border: 1px solid {THEME['border_color']};
        border-left: 4px solid {THEME['intro_border']};
        padding: 0.85rem 1.2rem;
        border-radius: 6px;
        margin-bottom: 1.25rem;
        font-size: 0.93rem;
        color: {THEME['intro_text']};
        line-height: 1.5;
    }}

    /* Company Header Card */
    .company-header-card {{
        background-color: {THEME['header_card_bg']};
        border: 1px solid {THEME['border_color']};
        border-radius: 10px;
        padding: 1.25rem 1.5rem;
        margin-bottom: 1.25rem;
        box-shadow: 0 1px 3px rgba(0,0,0,{"0.2" if is_dark else "0.03"});
    }}

    /* Badges */
    .company-badge {{
        background-color: {THEME['badge_aapl_bg']};
        color: {THEME['badge_aapl_text']};
        border: 1px solid {THEME['badge_aapl_border']};
        padding: 0.35rem 0.85rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        display: inline-block;
    }}
    .company-badge-secondary {{
        background-color: {THEME['badge_sec_bg']};
        color: {THEME['badge_sec_text']};
        border: 1px solid {THEME['badge_sec_border']};
        padding: 0.35rem 0.85rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        display: inline-block;
    }}

    /* Metric Cards */
    .metric-card {{
        background-color: {THEME['card_bg']};
        border: 1px solid {THEME['border_color']};
        border-radius: 8px;
        padding: 0.9rem 0.75rem;
        box-shadow: 0 1px 3px rgba(0,0,0,{"0.25" if is_dark else "0.04"});
        text-align: center;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }}
    .metric-category {{
        font-size: 0.72rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: {THEME['metric_cat_color']};
        margin-bottom: 0.2rem;
    }}
    .metric-name {{
        font-size: 0.88rem;
        font-weight: 600;
        color: {THEME['metric_name_color']};
        margin-bottom: 0.35rem;
    }}
    .metric-val {{
        font-size: 1.55rem;
        font-weight: 700;
        color: {THEME['metric_val_color']};
        line-height: 1.2;
    }}
    .metric-sub {{
        font-size: 0.74rem;
        color: {THEME['text_muted']};
        margin-top: 0.3rem;
        line-height: 1.2;
    }}
    .metric-sub.positive {{
        color: {THEME['positive']};
        font-weight: 600;
    }}
    .metric-sub.negative {{
        color: {THEME['negative']};
        font-weight: 600;
    }}
    .metric-sub-neutral {{
        font-size: 0.74rem;
        color: {THEME['text_secondary']};
        margin-top: 0.3rem;
        line-height: 1.2;
        font-weight: 500;
    }}

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {{
        border-bottom: 1px solid {THEME['border_color']};
        gap: 0.5rem;
    }}
    .stTabs [data-baseweb="tab"] {{
        color: {THEME['text_secondary']} !important;
        font-weight: 500;
        border-radius: 4px 4px 0 0;
    }}
    .stTabs [data-baseweb="tab"][aria-selected="true"] {{
        color: {THEME['accent_blue']} !important;
        font-weight: 700;
    }}
    .stTabs [data-baseweb="tab-highlight"] {{
        background-color: {THEME['accent_blue']} !important;
    }}

    /* Expanders */
    div[data-testid="stExpander"] {{
        background-color: {THEME['card_bg']};
        border: 1px solid {THEME['border_color']};
        border-radius: 8px;
    }}
    div[data-testid="stExpander"] summary {{
        color: {THEME['text_primary']} !important;
        font-weight: 600;
    }}

    /* Dataframe Container */
    [data-testid="stDataFrame"] {{
        border: 1px solid {THEME['border_color']};
        border-radius: 8px;
        background-color: {THEME['card_bg']};
    }}

    /* Captions */
    [data-testid="stCaptionContainer"] p, .stCaption {{
        color: {THEME['text_secondary']} !important;
    }}

    /* Info / Warning / Success Alert Boxes */
    div[data-testid="stAlert"] {{
        border-radius: 8px;
        border: 1px solid {THEME['border_color']};
    }}

    /* Form Controls & Inputs */
    div[data-baseweb="select"] > div,
    div[data-baseweb="base-input"],
    div[data-baseweb="input"] {{
        background-color: {THEME['card_secondary_bg']} !important;
        border-color: {THEME['border_color']} !important;
        color: {THEME['text_primary']} !important;
    }}
    div[data-baseweb="select"] span {{
        color: {THEME['text_primary']} !important;
    }}
    div[data-testid="stNumberInput"] input,
    div[data-testid="stTextInput"] input {{
        background-color: {THEME['card_secondary_bg']} !important;
        color: {THEME['text_primary']} !important;
        border-color: {THEME['border_color']} !important;
    }}
</style>
""", unsafe_allow_html=True)


# ==============================================================================
# HEADER & APPLICATION OVERVIEW
# ==============================================================================

st.markdown('<div class="main-title">SEC 10-K Financial & Sentiment Analyzer</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Automated SEC filing analysis, financial ratios, MD&A sentiment, stock prices & cross-company comparison</div>', unsafe_allow_html=True)

st.markdown("""
<div class="intro-box">
This executive dashboard retrieves SEC Form 10-K annual reports, extracts multi-year financial statements
and Item 7 MD&A narratives, executes Loughran-McDonald sentiment analysis, computes financial ratios,
retrieves historical stock prices, and provides interactive cross-company comparisons (Apple vs. Tesla).
</div>
""", unsafe_allow_html=True)


# ==============================================================================
# SESSION STATE INITIALIZATION
# ==============================================================================

if "analysis_cache" not in st.session_state:
    st.session_state.analysis_cache = {}

if "company_mode" not in st.session_state:
    st.session_state.company_mode = "Apple (AAPL)"

if "custom_ticker" not in st.session_state:
    st.session_state.custom_ticker = "AAPL"

if "active_ticker" not in st.session_state:
    st.session_state.active_ticker = "AAPL"

if "sidebar_ticker_input" not in st.session_state:
    st.session_state.sidebar_ticker_input = "AAPL"

if "selected_filing_count" not in st.session_state:
    st.session_state.selected_filing_count = 4


# ==============================================================================
# LOUGHRAN-MCDONALD DICTIONARY STATUS & DETECTION
# ==============================================================================

DEFAULT_BUNDLED_DICT_FILENAME = "Loughran-McDonald_MasterDictionary_1993-2025.csv"
dict_dir = setup_dictionaries_directory()

# Preserve original find_dictionary_file reference from main.py
if not hasattr(main, "_orig_find_dictionary_file"):
    main._orig_find_dictionary_file = main.find_dictionary_file

def get_active_dictionary():
    """
    Determine the active Loughran-McDonald dictionary for the current session.
    Prioritizes:
    1. Session-level replacement uploaded by user in the current session (st.session_state)
    2. Official bundled dictionary (dictionaries/Loughran-McDonald_MasterDictionary_1993-2025.csv)
    3. Any valid dictionary discovered in the dictionaries directory via main.py logic
    """
    # 1. Check for session-level custom override
    session_override_path = st.session_state.get("session_dict_override_path")
    if session_override_path and os.path.isfile(session_override_path):
        session_override_name = st.session_state.get(
            "session_dict_override_name",
            os.path.basename(session_override_path)
        )
        return session_override_path, session_override_name, True

    # 2. Check for default bundled official dictionary
    bundled_path = os.path.join(dict_dir, DEFAULT_BUNDLED_DICT_FILENAME)
    if os.path.isfile(bundled_path):
        return bundled_path, DEFAULT_BUNDLED_DICT_FILENAME, False

    # 3. Fallback discovery using existing main.py search
    found_path, found_name = main._orig_find_dictionary_file(dict_dir)
    return found_path, found_name, False

# Intercept main.find_dictionary_file so run_analysis_workflow transparently
# loads the active session dictionary without modifying backend step logic
def _session_find_dictionary_file(dictionaries_dir):
    path, name, _ = get_active_dictionary()
    if path and os.path.isfile(path):
        return path, name
    return main._orig_find_dictionary_file(dictionaries_dir)

main.find_dictionary_file = _session_find_dictionary_file

dict_filepath, dict_filename, is_custom_override = get_active_dictionary()
has_dictionary = dict_filepath is not None and os.path.isfile(dict_filepath)


# ==============================================================================
# HELPER: NEUTRAL YOY TREND INDICATOR
# ==============================================================================

def get_trend_indicator(curr_val, prev_val, is_pct=False, precision=2):
    """
    Generate neutral YoY trend indicator comparing current vs prior fiscal year.
    Returns: formatted string e.g. '↑ Increased (+0.12) vs prior yr'
    """
    if curr_val is None or prev_val is None:
        return "No prior comparison"
    diff = curr_val - prev_val
    threshold = max(0.005 * abs(prev_val), 1e-4) if prev_val != 0 else 1e-4
    if abs(diff) <= threshold:
        return "→ Approx. unchanged vs prior yr"
    elif diff > 0:
        if is_pct:
            return f"↑ Increased ({diff * 100:+.2f}%) vs prior yr"
        else:
            return f"↑ Increased ({diff:+.{precision}f}) vs prior yr"
    else:
        if is_pct:
            return f"↓ Decreased ({diff * 100:+.2f}%) vs prior yr"
        else:
            return f"↓ Decreased ({diff:+.{precision}f}) vs prior yr"


# ==============================================================================
# HELPER: STRICT HORIZON CACHE LOOKUP
# ==============================================================================

def get_cached_result(ticker, filing_count):
    """
    Retrieve cached analysis results strictly matching the requested filing-count horizon.
    """
    ticker_clean = str(ticker).strip().upper()
    # 1. Exact key match: (ticker, filing_count)
    res = st.session_state.analysis_cache.get((ticker_clean, filing_count))
    if res and len(res.get("analysis_data", [])) == filing_count:
        return res
    # 2. Fallback key match: ticker alone, only if the horizon matches
    fallback = st.session_state.analysis_cache.get(ticker_clean)
    if fallback and len(fallback.get("analysis_data", [])) == filing_count:
        return fallback
    return None


# ==============================================================================
# HELPER: UNIFIED PLOTLY CHART THEME
# ==============================================================================

def apply_chart_theme(
    fig,
    title="",
    xaxis_title="",
    yaxis_title="",
    yaxis2_title=None,
    height=380,
    barmode=None,
    show_legend=True,
    legend_top=True,
    margin=None,
    hovermode=None,
    tickprefix=None,
    ticksuffix=None,
    tickformat=None
):
    """
    Uniformly apply current THEME (Dark or Light) to a Plotly figure.
    """
    layout_dict = {
        "template": THEME["plotly_template"],
        "paper_bgcolor": THEME["chart_paper_bg"],
        "plot_bgcolor": THEME["chart_plot_bg"],
        "height": height,
        "font": dict(color=THEME["chart_font_color"]),
        "hoverlabel": dict(
            bgcolor=THEME["card_bg"],
            font_color=THEME["text_primary"],
            bordercolor=THEME["border_color"]
        )
    }

    if title:
        layout_dict["title"] = dict(
            text=f"<b>{title}</b>",
            font=dict(size=15, color=THEME["text_primary"] if is_dark else "#1F497D")
        )

    if barmode:
        layout_dict["barmode"] = barmode

    if margin:
        layout_dict["margin"] = margin

    if hovermode:
        layout_dict["hovermode"] = hovermode

    if show_legend:
        if legend_top:
            layout_dict["legend"] = dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
                font=dict(color=THEME["chart_font_color"])
            )
        else:
            layout_dict["legend"] = dict(
                font=dict(color=THEME["chart_font_color"])
            )
    else:
        layout_dict["showlegend"] = False

    fig.update_layout(**layout_dict)

    # Style X-axis
    fig.update_xaxes(
        title_text=xaxis_title,
        showgrid=True,
        gridcolor=THEME["chart_grid_color"],
        zeroline=False,
        title_font=dict(color=THEME["text_secondary"]),
        tickfont=dict(color=THEME["text_secondary"])
    )

    # Style primary Y-axis
    y_kwargs = {
        "title_text": yaxis_title,
        "showgrid": True,
        "gridcolor": THEME["chart_grid_color"],
        "zeroline": True,
        "zerolinecolor": THEME["chart_zeroline_color"],
        "title_font": dict(color=THEME["text_secondary"]),
        "tickfont": dict(color=THEME["text_secondary"]),
    }
    if tickprefix:
        y_kwargs["tickprefix"] = tickprefix
    if ticksuffix:
        y_kwargs["ticksuffix"] = ticksuffix
    if tickformat:
        y_kwargs["tickformat"] = tickformat
    if yaxis2_title:
        y_kwargs["secondary_y"] = False

    fig.update_yaxes(**y_kwargs)

    # Style secondary Y-axis if present
    if yaxis2_title:
        fig.update_yaxes(
            title_text=yaxis2_title,
            showgrid=False,
            title_font=dict(color=THEME["text_secondary"]),
            tickfont=dict(color=THEME["text_secondary"]),
            secondary_y=True
        )

    return fig


# ==============================================================================
# SIDEBAR / USER INPUT CONTROLS
# ==============================================================================

with st.sidebar:
    # --------------------------------------------------------------------------
    # THEME SELECTOR (PERSISTENT VIA st.session_state)
    # --------------------------------------------------------------------------
    st.subheader("🎨 Appearance")
    theme_options = ["🌙 Dark", "☀️ Light"]
    selected_theme = st.pills(
        "Theme:",
        options=theme_options,
        default=st.session_state.theme_mode,
        key="theme_mode_selector"
    )

    if selected_theme and selected_theme != st.session_state.theme_mode:
        st.session_state.theme_mode = selected_theme
        st.rerun()

    st.markdown("---")
    st.header("⚙️ Analysis Settings")
    st.markdown("Configure target company and historical horizon to run:")

    # Quick Switcher for Run Target
    company_options = ["Apple (AAPL)", "Tesla (TSLA)", "Custom Ticker"]
    selected_mode = st.pills(
        "Run Target:",
        options=company_options,
        default=st.session_state.company_mode,
        key="company_pill_selector"
    )

    if selected_mode and selected_mode != st.session_state.company_mode:
        st.session_state.company_mode = selected_mode
        if selected_mode == "Apple (AAPL)":
            st.session_state.active_ticker = "AAPL"
            st.session_state.sidebar_ticker_input = "AAPL"
        elif selected_mode == "Tesla (TSLA)":
            st.session_state.active_ticker = "TSLA"
            st.session_state.sidebar_ticker_input = "TSLA"
        elif selected_mode == "Custom Ticker":
            custom_val = st.session_state.get("custom_ticker", "AAPL")
            st.session_state.active_ticker = custom_val
            st.session_state.sidebar_ticker_input = custom_val
        st.rerun()

    # Ensure sidebar ticker widget state strictly reflects active selection
    if st.session_state.company_mode == "Apple (AAPL)":
        st.session_state.sidebar_ticker_input = "AAPL"
        st.session_state.active_ticker = "AAPL"
        is_custom = False
        ticker_help = "Selected via quick-switcher (Apple Inc. — AAPL)."
    elif st.session_state.company_mode == "Tesla (TSLA)":
        st.session_state.sidebar_ticker_input = "TSLA"
        st.session_state.active_ticker = "TSLA"
        is_custom = False
        ticker_help = "Selected via quick-switcher (Tesla, Inc. — TSLA)."
    else:
        if st.session_state.get("sidebar_ticker_input") in ("AAPL", "TSLA") and "custom_ticker" in st.session_state:
            st.session_state.sidebar_ticker_input = st.session_state.custom_ticker
            st.session_state.active_ticker = st.session_state.custom_ticker
        is_custom = True
        ticker_help = "Enter any public company ticker symbol listed on a U.S. exchange (e.g., MSFT, NVDA, AMZN)."

    input_ticker = st.text_input(
        "Ticker Symbol:",
        max_chars=10,
        disabled=not is_custom,
        help=ticker_help,
        key="sidebar_ticker_input"
    ).strip().upper()

    if is_custom:
        st.session_state.custom_ticker = input_ticker
        st.session_state.active_ticker = input_ticker
    else:
        st.session_state.active_ticker = st.session_state.sidebar_ticker_input

    input_filing_count = st.number_input(
        "Number of Recent 10-K Filings:",
        min_value=1,
        max_value=10,
        value=int(st.session_state.get("selected_filing_count", 4)),
        step=1,
        help="Select between 1 and 10 fiscal years (default is 4; final assignment uses 10).",
        key="sidebar_filing_count_input"
    )
    st.session_state.selected_filing_count = input_filing_count

    run_btn = st.button("🚀 Run Analysis", type="primary", width="stretch")

    st.markdown("---")
    st.subheader("📚 Sentiment Dictionary")
    if has_dictionary:
        status_label = f"**Loaded:** `{dict_filename}`"
        if is_custom_override:
            status_label += " *(Session Override)*"
        st.success(status_label)
        with st.expander("Replace / Upload New Dictionary", expanded=False):
            st.caption("Upload an updated Loughran-McDonald dictionary (.csv / .xlsx):")
            uploaded_sidebar = st.file_uploader(
                "Upload replacement dictionary",
                type=["csv", "xlsx"],
                key="sidebar_dict_uploader",
                label_visibility="collapsed"
            )
            if uploaded_sidebar is not None:
                upload_subfolder = os.path.join(dict_dir, ".session_uploads")
                os.makedirs(upload_subfolder, exist_ok=True)
                save_path = os.path.join(upload_subfolder, uploaded_sidebar.name)
                with open(save_path, "wb") as f_out:
                    f_out.write(uploaded_sidebar.getbuffer())
                check = load_loughran_mcdonald_dictionary(save_path)
                if check and check.get("total_words", 0) > 0:
                    st.session_state["session_dict_override_path"] = save_path
                    st.session_state["session_dict_override_name"] = uploaded_sidebar.name
                    st.success(f"Loaded {uploaded_sidebar.name} ({check['total_words']:,d} words)!")
                    st.rerun()
                else:
                    if os.path.exists(save_path):
                        os.remove(save_path)
                    st.error("Invalid dictionary file format. Please upload a .csv or .xlsx file.")
            if is_custom_override:
                if st.button("↺ Revert to Default Dictionary", key="reset_dict_btn", width="stretch"):
                    st.session_state.pop("session_dict_override_path", None)
                    st.session_state.pop("session_dict_override_name", None)
                    st.rerun()
    else:
        st.warning("⚠️ **Dictionary Missing**\n\nUpload required for Step 7 sentiment analysis.")

    st.markdown("---")
    st.caption("🔒 **SEC Fair Access Compliant**: Requests are throttled and use declarative User-Agent contact credentials.")
    st.caption("📚 **Accounting Standards**: Standard financial-analysis ratios calculated using U.S. GAAP financial-statement data.")


# ==============================================================================
# DICTIONARY UPLOAD FALLBACK INTERFACE
# ==============================================================================

if not has_dictionary:
    st.warning("### ⚠️ Loughran-McDonald Sentiment Dictionary Required")
    st.markdown("""
    To perform **Step 7 (MD&A Sentiment Analysis)**, this application requires the official **Loughran-McDonald Financial Sentiment Dictionary**.
    Please upload the official dictionary file (.csv or .xlsx) from Notre Dame SRAF below.
    """)

    uploaded_main_dict = st.file_uploader(
        "Upload Loughran-McDonald Master Dictionary (.csv or .xlsx):",
        type=["csv", "xlsx"],
        key="main_dict_uploader"
    )

    if uploaded_main_dict is not None:
        upload_subfolder = os.path.join(dict_dir, ".session_uploads")
        os.makedirs(upload_subfolder, exist_ok=True)
        save_path = os.path.join(upload_subfolder, uploaded_main_dict.name)
        try:
            with open(save_path, "wb") as f_out:
                f_out.write(uploaded_main_dict.getbuffer())

            check = load_loughran_mcdonald_dictionary(save_path)
            if check and check.get("total_words", 0) > 0:
                st.session_state["session_dict_override_path"] = save_path
                st.session_state["session_dict_override_name"] = uploaded_main_dict.name
                st.success(f"✅ Loaded {uploaded_main_dict.name} ({check['total_words']:,d} words)!")
                st.rerun()
            else:
                if os.path.exists(save_path):
                    os.remove(save_path)
                st.error("❌ Invalid dictionary file. Please upload an official CSV or Excel file.")
        except Exception as upload_err:
            if os.path.exists(save_path):
                os.remove(save_path)
            st.error(f"❌ Error saving dictionary: {upload_err}")


# ==============================================================================
# RUN ANALYSIS TRIGGER (RERUN-SAFE)
# ==============================================================================

if run_btn:
    clean_ticker = st.session_state.active_ticker.strip().upper()
    if not clean_ticker or not re.match(r"^[A-Z0-9\.\-]+$", clean_ticker):
        st.error("❌ Please enter a valid stock ticker symbol (letters, numbers, dots, or hyphens only).")
    elif not has_dictionary:
        st.error("❌ Cannot start analysis: The Loughran-McDonald sentiment dictionary is required. Please upload the dictionary file first.")
    else:
        status_box = st.status(f"Starting SEC 10-K Analysis for **{clean_ticker}** ({input_filing_count} filings)...", expanded=True)
        progress_bar = st.progress(0.0)

        def on_progress(step_num, stage_name, description):
            pct = min(1.0, step_num / 10.0)
            progress_bar.progress(pct)
            status_box.write(f"**Step {step_num}/10: {stage_name}** — {description}")

        try:
            workflow_results = run_analysis_workflow(
                ticker=clean_ticker,
                filing_count=input_filing_count,
                progress_callback=on_progress,
                is_cli=False
            )

            if workflow_results.get("success"):
                progress_bar.progress(1.0)
                status_box.update(label=f"✅ Analysis for {clean_ticker} complete!", state="complete", expanded=False)
                # Store in multi-company cache by exact (ticker, filing_count) and ticker
                cache_key = (clean_ticker, input_filing_count)
                st.session_state.analysis_cache[cache_key] = workflow_results
                st.session_state.analysis_cache[clean_ticker] = workflow_results
                st.rerun()
            else:
                progress_bar.empty()
                status_box.update(label=f"❌ Analysis failed for {clean_ticker}", state="error", expanded=True)
                st.error(f"**Error:** {workflow_results.get('error')}")

        except Exception as ex:
            progress_bar.empty()
            status_box.update(label="❌ Unexpected Error", state="error", expanded=True)
            st.error(f"An unexpected error occurred during execution: {ex}")


# ==============================================================================
# RENDER INDIVIDUAL COMPANY DASHBOARD (REUSABLE COMPONENT)
# ==============================================================================

def render_company_dashboard(results, target_filing_count, prefix="company"):
    company_name = results["company_name"]
    ticker = results["ticker"]
    cik = results["cik"]
    analysis_data = results["analysis_data"]
    num_years = len(analysis_data)
    statement_data = results["statement_data"]
    sentiment_results = results["sentiment_results"]
    stock_price_results = results["stock_price_results"]
    raw_accounts_info = results["raw_accounts_info"]
    available_cats = results["available_categories"]
    output_filepath = results["output_filepath"]

    oldest = analysis_data[0]
    latest = analysis_data[-1]
    prev = analysis_data[-2] if num_years >= 2 else None

    fiscal_horizon = f"FY {oldest['year']} – FY {latest['year']}"
    reporting_scope = f"{num_years}-Year SEC 10-K Analysis"
    cik_str = str(cik).zfill(10)

    # Polished Company Header
    st.markdown(f"""
    <div class="company-header-card">
        <div style="display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 0.85rem;">
            <div>
                <div style="font-size: 1.7rem; font-weight: 700; color: {THEME['text_primary']}; line-height: 1.25;">
                    🏢 {company_name} <span style="font-size: 1.2rem; color: {THEME['text_secondary']}; font-weight: 500;">({ticker})</span>
                </div>
                <div style="font-size: 0.92rem; color: {THEME['text_secondary']}; margin-top: 0.35rem; line-height: 1.4;">
                    SEC Central Index Key: <b>CIK {cik_str}</b> &nbsp;•&nbsp; Reporting Period: <b>{fiscal_horizon}</b> &nbsp;•&nbsp; Master Workbook: <b>{Path(output_filepath).name}</b>
                </div>
            </div>
            <div style="display: flex; gap: 0.5rem; align-items: center; flex-wrap: wrap;">
                <span class="company-badge">📋 {reporting_scope}</span>
                <span class="company-badge-secondary">📅 FY {oldest['year']} – FY {latest['year']}</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Seven Navigation Tabs for Company Dashboard
    tab_overview, tab_stock_sent, tab_statements, tab_ratios, tab_sentiment, tab_growth, tab_download = st.tabs([
        "📊 Overview",
        "📈 Stock & Sentiment",
        "📑 Financial Statements",
        "📉 10-Year Financial Ratios",
        "📝 Sentiment Analysis",
        "📊 Growth & Performance",
        "📥 Download Report"
    ])

    # --------------------------------------------------------------------------
    # TAB 1: OVERVIEW
    # --------------------------------------------------------------------------
    with tab_overview:
        # Macro Performance Metrics
        oldest_p = oldest.get("stock_price")
        latest_p = latest.get("stock_price")
        prev_p = prev.get("stock_price") if prev else None

        if oldest_p is not None and latest_p is not None and oldest_p > 0:
            stock_ret_pct = ((latest_p - oldest_p) / oldest_p) * 100.0
            stock_ret_str = f"{stock_ret_pct:+.1f}%"
            stock_ret_cls = "positive" if stock_ret_pct >= 0 else "negative"
            stock_ret_sub = f"FY {oldest['year']} (${oldest_p:,.2f}) → FY {latest['year']} (${latest_p:,.2f})"
        else:
            stock_ret_str = "N/A"
            stock_ret_cls = ""
            stock_ret_sub = "Price data unavailable"

        by_year = raw_accounts_info.get("by_year", {})
        oldest_rev = by_year.get(oldest["year"], {}).get("revenue")
        latest_rev = by_year.get(latest["year"], {}).get("revenue")
        prev_rev = by_year.get(prev["year"], {}).get("revenue") if prev else None

        if oldest_rev is not None and latest_rev is not None and oldest_rev > 0:
            rev_growth_pct = ((latest_rev - oldest_rev) / oldest_rev) * 100.0
            rev_growth_str = f"{rev_growth_pct:+.1f}%"
            rev_growth_cls = "positive" if rev_growth_pct >= 0 else "negative"
            rev_growth_sub = f"FY {oldest['year']} (${oldest_rev:,.0f}) → FY {latest['year']} (${latest_rev:,.0f})"
        else:
            rev_growth_str = "N/A"
            rev_growth_cls = ""
            rev_growth_sub = "Revenue data unavailable"

        latest_pn = latest.get("pn_ratio")
        prev_pn = prev.get("pn_ratio") if prev else None
        latest_pn_str = f"{latest_pn:.2f}" if latest_pn is not None else "N/A"
        latest_pn_trend = get_trend_indicator(latest_pn, prev_pn, precision=2)

        latest_sp_str = f"${latest_p:,.2f}" if latest_p is not None else "N/A"
        latest_sp_trend = get_trend_indicator(latest_p, prev_p, precision=2)

        st.subheader("📌 Executive Summary")
        st.caption(f"Comprehensive executive snapshot for **{company_name}** across the {reporting_scope.lower()}.")

        st.markdown(f"**Period Macro Performance ({reporting_scope}: FY {oldest['year']} – FY {latest['year']})**")
        c_p1, c_p2, c_p3, c_p4 = st.columns(4)

        c_p1.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Market Performance</div>
            <div class="metric-name">Period Stock Return</div>
            <div class="metric-val">{stock_ret_str}</div>
            <div class="metric-sub {stock_ret_cls}">{stock_ret_sub}</div>
        </div>
        """, unsafe_allow_html=True)

        c_p2.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Top-Line Expansion</div>
            <div class="metric-name">Period Revenue Growth</div>
            <div class="metric-val">{rev_growth_str}</div>
            <div class="metric-sub {rev_growth_cls}">{rev_growth_sub}</div>
        </div>
        """, unsafe_allow_html=True)

        c_p3.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Market Value</div>
            <div class="metric-name">Latest Stock Price</div>
            <div class="metric-val">{latest_sp_str}</div>
            <div class="metric-sub-neutral">{latest_sp_trend}</div>
        </div>
        """, unsafe_allow_html=True)

        c_p4.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">MD&A Sentiment</div>
            <div class="metric-name">Latest P/N Ratio</div>
            <div class="metric-val">{latest_pn_str}</div>
            <div class="metric-sub-neutral">{latest_pn_trend}</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"**Latest Fiscal Year Accounting Health (FY {latest['year']})**")
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)

        # 1. Current Ratio
        cr_val = f"{latest['current_ratio']:.2f}" if latest.get("current_ratio") is not None else "N/A"
        cr_trend = get_trend_indicator(latest.get("current_ratio"), prev.get("current_ratio") if prev else None, precision=2)
        col_m1.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Liquidity</div>
            <div class="metric-name">Current Ratio</div>
            <div class="metric-val">{cr_val}</div>
            <div class="metric-sub-neutral">{cr_trend}</div>
        </div>
        """, unsafe_allow_html=True)

        # 2. Debt-to-Assets
        da_val = f"{latest['debt_to_assets'] * 100:.2f}%" if latest.get("debt_to_assets") is not None else "N/A"
        da_trend = get_trend_indicator(latest.get("debt_to_assets"), prev.get("debt_to_assets") if prev else None, is_pct=True)
        col_m2.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Solvency</div>
            <div class="metric-name">Debt-to-Assets</div>
            <div class="metric-val">{da_val}</div>
            <div class="metric-sub-neutral">{da_trend}</div>
        </div>
        """, unsafe_allow_html=True)

        # 3. Total Asset Turnover
        at_val = f"{latest['asset_turnover']:.2f}" if latest.get("asset_turnover") is not None else "N/A"
        at_trend = get_trend_indicator(latest.get("asset_turnover"), prev.get("asset_turnover") if prev else None, precision=2)
        col_m3.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Efficiency</div>
            <div class="metric-name">Asset Turnover</div>
            <div class="metric-val">{at_val}</div>
            <div class="metric-sub-neutral">{at_trend}</div>
        </div>
        """, unsafe_allow_html=True)

        # 4. Net Profit Margin
        pm_val = f"{latest['profit_margin'] * 100:.2f}%" if latest.get("profit_margin") is not None else "N/A"
        pm_trend = get_trend_indicator(latest.get("profit_margin"), prev.get("profit_margin") if prev else None, is_pct=True)
        col_m4.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Profitability</div>
            <div class="metric-name">Net Margin</div>
            <div class="metric-val">{pm_val}</div>
            <div class="metric-sub-neutral">{pm_trend}</div>
        </div>
        """, unsafe_allow_html=True)

        # Historical Summary Table
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("📋 Historical Summary")
        st.caption(f"Chronological financial ratios, stock prices, and MD&A sentiment scores (oldest to newest: FY {oldest['year']} to FY {latest['year']}).")

        hist_summary_rows = []
        for r in analysis_data:
            sp = f"${r['stock_price']:,.2f}" if r.get("stock_price") is not None else "N/A"
            cr = f"{r['current_ratio']:.2f}" if r.get("current_ratio") is not None else "N/A"
            da = f"{r['debt_to_assets'] * 100:.2f}%" if r.get("debt_to_assets") is not None else "N/A"
            at = f"{r['asset_turnover']:.2f}" if r.get("asset_turnover") is not None else "N/A"
            pm = f"{r['profit_margin'] * 100:.2f}%" if r.get("profit_margin") is not None else "N/A"
            pn = f"{r['pn_ratio']:.2f}" if r.get("pn_ratio") is not None else "N/A"
            hist_summary_rows.append({
                "Fiscal Year": f"FY {r['year']}",
                "Stock Price": sp,
                "Current Ratio": cr,
                "Debt-to-Assets": da,
                "Total Asset Turnover": at,
                "Net Profit Margin": pm,
                "P/N Ratio": pn
            })
        st.dataframe(pd.DataFrame(hist_summary_rows), width="stretch", hide_index=True)

        # Data Sources & Methodology Expander
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("📖 Data Sources & Methodology", expanded=False):
            st.markdown("""
            ### 🏛️ 1. SEC EDGAR Data Source & Filing Retrieval
            - **Primary Source**: U.S. Securities and Exchange Commission (SEC EDGAR) official submissions API (`https://data.sec.gov/submissions/`).
            - **Filing Form**: Form 10-K (Annual Reports filed by public entities under Section 13 or 15(d) of the Securities Exchange Act of 1934).
            - **Entity Lookup**: Official SEC Central Index Key (CIK) mapped directly from the company's exchange ticker symbol.
            - **Fair Access Compliance**: Declarative User-Agent contact header, polite pacing, and request rate throttling adhering strictly to SEC Fair Access regulations.

            ---
            ### 📅 2. Fiscal Year & `reportDate` Alignment
            - **Period End Alignment**: Each filing's financial statements are keyed to the official period-end date (`reportDate` / balance sheet date).
            - **Multi-Year Comparative Reporting**: Financial statements are aligned chronologically from oldest to newest fiscal year.
            - **Account Mapping**: Standard U.S. GAAP taxonomy concepts are extracted directly from primary financial statements.

            ---
            ### 📈 3. Stock Price Methodology & Previous-Trading-Day Rule
            - **Source**: Yahoo Finance via `yfinance`.
            - **Price Metric**: Historical split-adjusted closing price (`Close`).
            - **Alignment Rule**: Stock prices are matched to the exact fiscal year-end balance sheet date.
            - **Non-Trading Day Rule**: When the balance sheet date falls on a weekend or market holiday, the system automatically selects the **most recent previous trading day**. Future trading dates are never used.

            ---
            ### 📝 4. Loughran-McDonald Financial Sentiment Scoring
            - **Source**: University of Notre Dame Software Repository for Accounting and Finance (SRAF), developed by Tim Loughran and Bill McDonald.
            - **Target Text**: Item 7 — *Management's Discussion and Analysis of Financial Condition and Results of Operations (MD&A)*.
            - **Text Preprocessing**: Raw HTML/XBRL tags and tables stripped; punctuation and possessive suffixes normalized; tokenized into uppercase alphabetic words without stemming.
            - **Dictionary Categories**: 8 domain-specific categories (*Positive, Negative, Uncertainty, Litigious, Strong Modal, Weak Modal, Constraining, Complexity*).
            - **Positive-to-Negative (P/N) Ratio**:
              $$\\text{P/N Ratio} = \\frac{\\text{Positive Words}}{\\text{Negative Words}}$$

            ---
            ### 🔢 5. Financial Ratio Formulas & Definitions
            *Standard financial-analysis ratios calculated using U.S. GAAP financial-statement data.*
            - **Current Ratio (Liquidity)**: `Current Assets / Current Liabilities`
            - **Debt-to-Assets Ratio (Solvency)**: `Total Liabilities / Total Assets`
            - **Total Asset Turnover (Efficiency)**: `Revenue / Average Total Assets`
            - **Net Profit Margin (Profitability)**: `Net Income / Revenue`
            """)

    # --------------------------------------------------------------------------
    # TAB 2: STOCK & SENTIMENT
    # --------------------------------------------------------------------------
    with tab_stock_sent:
        st.subheader("📈 Stock Price vs. MD&A Sentiment P/N Ratio")
        st.caption("Chronological time-series comparison with dual Y-axes (Stock Price on Left, P/N Ratio on Right).")

        chrono_years = [str(r["year"]) for r in analysis_data]
        stock_series = [r.get("stock_price") for r in analysis_data]
        pn_series = [r.get("pn_ratio") for r in analysis_data]

        fig_dual = make_subplots(specs=[[{"secondary_y": True}]])
        fig_dual.add_trace(
            go.Scatter(
                x=chrono_years,
                y=stock_series,
                name="Stock Price ($)",
                mode="lines+markers",
                line=dict(color=THEME["chart_primary_line"], width=3),
                marker=dict(size=8, symbol="circle"),
                hovertemplate="FY %{x}<br>Stock Price: $%{y:,.2f}<extra></extra>"
            ),
            secondary_y=False
        )
        fig_dual.add_trace(
            go.Scatter(
                x=chrono_years,
                y=pn_series,
                name="P/N Ratio",
                mode="lines+markers",
                line=dict(color=THEME["chart_secondary_line"], width=3, dash="dot"),
                marker=dict(size=8, symbol="diamond"),
                hovertemplate="FY %{x}<br>P/N Ratio: %{y:.2f}<extra></extra>"
            ),
            secondary_y=True
        )
        apply_chart_theme(
            fig_dual,
            title=f"{company_name} ({ticker}) — Stock Price vs. MD&A Sentiment P/N Ratio",
            xaxis_title="Fiscal Year",
            yaxis_title="Historical Stock Price ($)",
            yaxis2_title="MD&A Sentiment P/N Ratio",
            height=460,
            hovermode="x unified",
            margin=dict(l=60, r=60, t=80, b=50),
            tickprefix="$"
        )
        st.plotly_chart(fig_dual, width="stretch", key=f"{prefix}_{ticker}_dual_axis")

        # Annual Stock Return Chart (New Feature)
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("📊 Annual Stock Price Return (% YoY)")
        st.caption("Year-over-year percentage price change calculated from official fiscal year-end closing prices (oldest baseline year excluded).")

        ret_years = []
        ret_vals = []
        ret_meta = []
        for i in range(1, len(analysis_data)):
            p_p = analysis_data[i-1].get("stock_price")
            p_c = analysis_data[i].get("stock_price")
            y_c = analysis_data[i]["year"]
            y_p = analysis_data[i-1]["year"]
            if p_p and p_c and p_p > 0:
                ret = ((p_c / p_p) - 1.0) * 100.0
                ret_years.append(f"FY {y_c}")
                ret_vals.append(ret)
                ret_meta.append((p_c, p_p, y_p))

        if ret_years:
            bar_colors = [THEME["chart_positive_color"] if v >= 0 else THEME["chart_negative_color"] for v in ret_vals]
            fig_ret = go.Figure(go.Bar(
                x=ret_years,
                y=ret_vals,
                marker_color=bar_colors,
                customdata=ret_meta,
                hovertemplate="<b>%{x}</b><br>YoY Return: %{y:+.2f}%<br>Price: $%{customdata[0]:,.2f} (FY %{customdata[2]}: $%{customdata[1]:,.2f})<extra></extra>"
            ))
            apply_chart_theme(
                fig_ret,
                title=f"{company_name} ({ticker}) — Annual Stock Price Return",
                xaxis_title="Fiscal Year",
                yaxis_title="Annual Return (%)",
                height=360,
                show_legend=False,
                margin=dict(l=60, r=40, t=60, b=40),
                ticksuffix="%"
            )
            st.plotly_chart(fig_ret, width="stretch", key=f"{prefix}_{ticker}_annual_ret")

        # Exploratory Correlation Expander
        corr_pairs = []
        for i in range(1, len(analysis_data)):
            p_prev = analysis_data[i-1].get("stock_price")
            p_curr = analysis_data[i].get("stock_price")
            pn_curr = analysis_data[i].get("pn_ratio")
            y_curr = analysis_data[i]["year"]
            y_prev = analysis_data[i-1]["year"]
            if p_prev and p_curr and pn_curr is not None and p_prev > 0:
                ret = (p_curr / p_prev) - 1.0
                corr_pairs.append({
                    "Fiscal Year": f"FY {y_curr}",
                    "Return Description": f"FY {y_curr} vs. FY {y_prev}",
                    "Stock Return": ret,
                    "Stock Return Pct": ret * 100.0,
                    "P/N Ratio": pn_curr
                })

        n_obs = len(corr_pairs)
        corr_val = None
        if n_obs >= 3:
            s_ret = pd.Series([cp["Stock Return"] for cp in corr_pairs])
            s_pn = pd.Series([cp["P/N Ratio"] for cp in corr_pairs])
            if s_ret.std() > 0 and s_pn.std() > 0:
                corr_val = s_ret.corr(s_pn)

        with st.expander("📊 Exploratory Correlation: Annual Stock Price Return vs. MD&A P/N Ratio", expanded=False):
            st.markdown("""
            This exploratory section calculates the Pearson statistical correlation ($r$) between:
            - **Annual Stock-Price Return**: $\\left(\\frac{\\text{Stock Price}_t}{\\text{Stock Price}_{t-1}}\\right) - 1$
            - **Current Fiscal Year MD&A P/N Ratio**: $(\\text{P/N Ratio})_t$
            """)

            if corr_val is not None:
                m_c1, m_c2 = st.columns([1, 2])
                m_c1.metric(
                    label="Pearson Correlation (r)",
                    value=f"{corr_val:+.3f}",
                    help="Pearson correlation coefficient ranging from -1.0 (inverse) to +1.0 (positive)."
                )
                caution_text = ""
                if n_obs < 4:
                    caution_text = "\n\n⚠️ *Limited sample — interpret cautiously* ($n < 4$)."

                m_c2.markdown(f"""
                - **Paired Observations ($n$):** `{n_obs}`  
                - **Observation Window:** {corr_pairs[0]['Fiscal Year']} to {corr_pairs[-1]['Fiscal Year']}{caution_text}
                """)
            else:
                st.info(f"Insufficient annual historical periods to calculate correlation (minimum 3 required; currently {n_obs} available).")

            st.warning("""
            ⚠️ **Correlation does not imply causation.**  
            This statistical metric is provided solely for academic exploratory analysis. Statistical correlation does not establish that MD&A sentiment causes or predicts stock-price movement, nor does it constitute an investment recommendation or financial advice.
            """)

            if corr_pairs:
                df_corr = pd.DataFrame([
                    {
                        "Fiscal Year": cp["Fiscal Year"],
                        "Period Evaluated": cp["Return Description"],
                        "Annual Stock Return": f"{cp['Stock Return Pct']:+.2f}%",
                        "MD&A P/N Ratio": f"{cp['P/N Ratio']:.2f}"
                    }
                    for cp in corr_pairs
                ])
                st.dataframe(df_corr, width="stretch", hide_index=True)

    # --------------------------------------------------------------------------
    # TAB 3: FINANCIAL STATEMENTS
    # --------------------------------------------------------------------------
    with tab_statements:
        st.subheader("📑 Primary Financial Statements (Item 8)")
        st.caption("Financial statements extracted from Item 8 of the company's SEC Form 10-K filings.")

        # Sub-tabs for financial statements
        sub_inc, sub_bal, sub_cf = st.tabs(["Income Statement", "Balance Sheet", "Cash Flows"])

        for stmt_name, sub_tab in [("Income Statement", sub_inc), ("Balance Sheet", sub_bal), ("Cash Flows", sub_cf)]:
            with sub_tab:
                stmt_obj = statement_data.get(stmt_name, {})
                years = stmt_obj.get("years", [])
                unit = stmt_obj.get("unit", "")
                items = stmt_obj.get("items", [])
                values = stmt_obj.get("values", {})

                if unit:
                    st.caption(f"**Reporting Scale / Units:** *{unit}*")

                if items and years:
                    stmt_rows = []
                    for item in items:
                        key, label, is_header = item
                        row_dict = {"Line Item": label.upper() if is_header else f"    {label}"}
                        row_vals = values.get(key, {})
                        row_text = f"{key[0]} {label}".lower()
                        for yr in years:
                            val = row_vals.get(yr)
                            if is_header or val is None:
                                row_dict[yr] = ""
                            elif "shares" in row_text:
                                row_dict[yr] = f"{val:,.0f}"
                            elif "per share" in row_text:
                                row_dict[yr] = f"${val:,.2f}" if val >= 0 else f"(${-val:,.2f})"
                            else:
                                row_dict[yr] = f"${val:,.0f}" if val >= 0 else f"(${-val:,.0f})"
                        stmt_rows.append(row_dict)

                    df_stmt = pd.DataFrame(stmt_rows)
                    st.dataframe(df_stmt, width="stretch", hide_index=True)
                else:
                    st.warning(f"No statement data available for {stmt_name}.")

    # --------------------------------------------------------------------------
    # TAB 4: 10-YEAR FINANCIAL RATIOS
    # --------------------------------------------------------------------------
    with tab_ratios:
        st.subheader("📉 10-Year Financial Ratio Trajectories")
        st.caption(f"Multi-year financial ratio trend visualization for **{company_name}**.")

        chrono_years = [str(r["year"]) for r in analysis_data]
        cr_series = [r.get("current_ratio") for r in analysis_data]
        da_series = [(r.get("debt_to_assets") * 100.0) if r.get("debt_to_assets") is not None else None for r in analysis_data]
        at_series = [r.get("asset_turnover") for r in analysis_data]
        pm_series = [(r.get("profit_margin") * 100.0) if r.get("profit_margin") is not None else None for r in analysis_data]

        # Four Plotly Trend Charts (2x2 Grid)
        c_r1, c_r2 = st.columns(2)
        with c_r1:
            fig_cr = go.Figure(go.Scatter(
                x=chrono_years, y=cr_series, mode="lines+markers",
                line=dict(color=THEME["chart_primary_line"], width=2.5), marker=dict(size=7),
                hovertemplate="FY %{x}<br>Current Ratio: %{y:.2f}<extra></extra>"
            ))
            apply_chart_theme(
                fig_cr,
                title="1. Current Ratio — Liquidity",
                xaxis_title="Fiscal Year",
                yaxis_title="Current Ratio",
                height=300,
                show_legend=False,
                margin=dict(l=40, r=20, t=50, b=40)
            )
            st.plotly_chart(fig_cr, width="stretch", key=f"{prefix}_{ticker}_cr_trend")

        with c_r2:
            fig_da = go.Figure(go.Scatter(
                x=chrono_years, y=da_series, mode="lines+markers",
                line=dict(color=THEME["chart_negative_color"], width=2.5), marker=dict(size=7),
                hovertemplate="FY %{x}<br>Debt-to-Assets: %{y:.2f}%<extra></extra>"
            ))
            apply_chart_theme(
                fig_da,
                title="2. Debt-to-Assets Ratio — Solvency",
                xaxis_title="Fiscal Year",
                yaxis_title="Debt-to-Assets (%)",
                height=300,
                show_legend=False,
                margin=dict(l=40, r=20, t=50, b=40),
                ticksuffix="%"
            )
            st.plotly_chart(fig_da, width="stretch", key=f"{prefix}_{ticker}_da_trend")

        c_r3, c_r4 = st.columns(2)
        with c_r3:
            fig_at = go.Figure(go.Scatter(
                x=chrono_years, y=at_series, mode="lines+markers",
                line=dict(color=THEME["accent_blue"], width=2.5), marker=dict(size=7),
                hovertemplate="FY %{x}<br>Asset Turnover: %{y:.2f}<extra></extra>"
            ))
            apply_chart_theme(
                fig_at,
                title="3. Total Asset Turnover — Efficiency",
                xaxis_title="Fiscal Year",
                yaxis_title="Asset Turnover",
                height=300,
                show_legend=False,
                margin=dict(l=40, r=20, t=50, b=40)
            )
            st.plotly_chart(fig_at, width="stretch", key=f"{prefix}_{ticker}_at_trend")

        with c_r4:
            fig_pm = go.Figure(go.Scatter(
                x=chrono_years, y=pm_series, mode="lines+markers",
                line=dict(color=THEME["chart_positive_color"], width=2.5), marker=dict(size=7),
                hovertemplate="FY %{x}<br>Net Profit Margin: %{y:.2f}%<extra></extra>"
            ))
            apply_chart_theme(
                fig_pm,
                title="4. Net Profit Margin — Profitability",
                xaxis_title="Fiscal Year",
                yaxis_title="Net Margin (%)",
                height=300,
                show_legend=False,
                margin=dict(l=40, r=20, t=50, b=40),
                ticksuffix="%"
            )
            st.plotly_chart(fig_pm, width="stretch", key=f"{prefix}_{ticker}_pm_trend")

        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("📋 Ratio Calculation Summary Table")
        ratio_table_data = []
        for r in analysis_data:
            ratio_table_data.append({
                "Fiscal Year": f"FY {r['year']}",
                "Current Ratio (Liquidity)": f"{r['current_ratio']:.2f}" if r["current_ratio"] is not None else "N/A",
                "Debt-to-Assets Ratio (Solvency)": f"{r['debt_to_assets'] * 100:.2f}%" if r["debt_to_assets"] is not None else "N/A",
                "Total Asset Turnover (Efficiency)": f"{r['asset_turnover']:.2f}" if r["asset_turnover"] is not None else "N/A",
                "Net Profit Margin (Profitability)": f"{r['profit_margin'] * 100:.2f}%" if r["profit_margin"] is not None else "N/A"
            })
        st.dataframe(pd.DataFrame(ratio_table_data), width="stretch", hide_index=True)

        st.markdown("""
        #### 📖 Ratio Formulas & Accounting Definitions
        - **Current Ratio (Liquidity)**: `Current Assets / Current Liabilities`
        - **Debt-to-Assets Ratio (Solvency)**: `Total Liabilities / Total Assets`
        - **Total Asset Turnover (Efficiency)**: `Revenue / Average Total Assets`
        - **Net Profit Margin (Profitability)**: `Net Income / Revenue`
        """)

        with st.expander("🔍 View Underlying Financial Statement Accounts (Audit Verification)"):
            by_year = raw_accounts_info.get("by_year", {})
            accts_data = []
            for r in analysis_data:
                yr = r["year"]
                item = by_year.get(yr, {})
                accts_data.append({
                    "Fiscal Year": f"FY {yr}",
                    "Revenue": f"${item.get('revenue'):,.0f}" if item.get("revenue") is not None else "N/A",
                    "Net Income": f"${item.get('net_income'):,.0f}" if item.get("net_income") is not None else "N/A",
                    "Current Assets": f"${item.get('current_assets'):,.0f}" if item.get("current_assets") is not None else "N/A",
                    "Current Liabilities": f"${item.get('current_liabilities'):,.0f}" if item.get("current_liabilities") is not None else "N/A",
                    "Total Liabilities": f"${item.get('total_liabilities'):,.0f}" if item.get("total_liabilities") is not None else "N/A",
                    "Ending Assets": f"${item.get('ending_assets'):,.0f}" if item.get("ending_assets") is not None else "N/A",
                    "Beginning Assets": f"${item.get('beginning_assets'):,.0f}" if item.get("beginning_assets") is not None else "N/A",
                    "Average Assets": f"${item.get('average_assets'):,.0f}" if item.get("average_assets") is not None else "N/A"
                })
            st.dataframe(pd.DataFrame(accts_data), width="stretch", hide_index=True)

    # --------------------------------------------------------------------------
    # TAB 5: SENTIMENT ANALYSIS
    # --------------------------------------------------------------------------
    with tab_sentiment:
        st.subheader(f"📝 MD&A Sentiment Analysis — {company_name} ({ticker})")
        st.caption("Item 7 Management's Discussion and Analysis (MD&A) narratives scored against the official Loughran-McDonald Master Sentiment Dictionary.")

        chrono_years = [f"FY {r['year']}" for r in analysis_data]
        pos_words = [r.get("positive", 0) for r in analysis_data]
        neg_words = [r.get("negative", 0) for r in analysis_data]
        pn_ratios = [r.get("pn_ratio") for r in analysis_data]

        # 1. POSITIVE VS. NEGATIVE WORD COUNTS
        fig_pn_bars = go.Figure()
        fig_pn_bars.add_trace(go.Bar(
            x=chrono_years,
            y=pos_words,
            name="Positive Financial Words",
            marker_color=THEME["chart_primary_line"] if is_dark else "#2B6CB0",
            hovertemplate="<b>%{x}</b><br>Category: Positive Words<br>Word Count: %{y:,d}<extra></extra>"
        ))
        fig_pn_bars.add_trace(go.Bar(
            x=chrono_years,
            y=neg_words,
            name="Negative Financial Words",
            marker_color=THEME["chart_negative_color"] if is_dark else "#E4572E",
            hovertemplate="<b>%{x}</b><br>Category: Negative Words<br>Word Count: %{y:,d}<extra></extra>"
        ))
        apply_chart_theme(
            fig_pn_bars,
            title=f"Loughran-McDonald MD&A Sentiment Word Counts ({ticker})",
            xaxis_title="Fiscal Reporting Period",
            yaxis_title="Word Count (Item 7 MD&A)",
            barmode="group",
            height=380,
            margin=dict(l=60, r=30, t=70, b=50)
        )
        st.plotly_chart(fig_pn_bars, width="stretch", key=f"{prefix}_{ticker}_pn_bars")

        # 2. 10-YEAR SENTIMENT FREQUENCY DATASET
        st.markdown("<br>", unsafe_allow_html=True)
        freq_section_title = f"{num_years}-Year Sentiment Frequency Dataset" if num_years != 10 else "10-Year Sentiment Frequency Dataset"
        st.subheader(f"📋 {freq_section_title}")
        st.caption(f"Chronological MD&A word frequencies and positive/negative sentiment ratios across {num_years} fiscal periods (oldest to newest):")

        freq_table_rows = []
        for r in analysis_data:
            tot = r.get("total_words", 0)
            pos = r.get("positive", 0)
            neg = r.get("negative", 0)
            unc = r.get("other_sentiment", {}).get("Uncertainty")
            lit = r.get("other_sentiment", {}).get("Litigious")
            pn = r.get("pn_ratio")

            pos_freq = (pos / tot * 100.0) if tot and tot > 0 else 0.0
            neg_freq = (neg / tot * 100.0) if tot and tot > 0 else 0.0

            freq_table_rows.append({
                "Fiscal Year": f"FY {r['year']}",
                "Total Words": f"{tot:,d}" if tot is not None else "N/A",
                "Positive Count": f"{pos:,d}" if pos is not None else "N/A",
                "Negative Count": f"{neg:,d}" if neg is not None else "N/A",
                "Uncertainty Count": f"{unc:,d}" if unc is not None else "N/A",
                "Litigious Count": f"{lit:,d}" if lit is not None else "N/A",
                "P/N Ratio": f"{pn:.2f}" if pn is not None else "N/A",
                "Pos Freq (%)": f"{pos_freq:.2f}%",
                "Neg Freq (%)": f"{neg_freq:.2f}%"
            })
        st.dataframe(pd.DataFrame(freq_table_rows), width="stretch", hide_index=True)

        # 3. P/N RATIO TREND
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("📈 MD&A Positive-to-Negative Ratio")
        st.caption("Chronological trajectory of Item 7 MD&A positive-to-negative sentiment score relative to the neutral baseline (P/N = 1.0).")

        fig_pn_line = go.Figure()
        fig_pn_line.add_trace(go.Scatter(
            x=chrono_years,
            y=pn_ratios,
            mode="lines+markers",
            name="P/N Ratio",
            line=dict(color=THEME["chart_primary_line"], width=3),
            marker=dict(size=8, symbol="diamond"),
            hovertemplate="<b>%{x}</b><br>P/N Ratio: %{y:.2f}<extra></extra>"
        ))
        fig_pn_line.add_hline(
            y=1.0,
            line_dash="dash",
            line_color=THEME["chart_neutral_line"],
            line_width=1.5,
            annotation_text="P/N = 1.0 (Neutral Baseline)",
            annotation_position="top right",
            annotation_font_color=THEME["text_secondary"]
        )
        apply_chart_theme(
            fig_pn_line,
            title=f"MD&A Positive-to-Negative Ratio — {company_name} ({ticker})",
            xaxis_title="Fiscal Reporting Period",
            yaxis_title="P/N Ratio",
            height=360,
            show_legend=False,
            margin=dict(l=60, r=40, t=60, b=50),
            tickformat=".2f"
        )
        st.plotly_chart(fig_pn_line, width="stretch", key=f"{prefix}_{ticker}_pn_line")

        st.info("""
        **P/N Ratio Interpretation**:  
        P/N > 1 means positive Loughran-McDonald word matches exceeded negative word matches. P/N < 1 means negative word matches exceeded positive word matches.
        """)

        # 4. DETAILED SENTIMENT DATA (Expander)
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("🔍 View Detailed Sentiment Categories", expanded=False):
            st.caption("Complete breakdown of word counts across all 8 domain-specific Loughran-McDonald sentiment categories (ordered chronologically from oldest to newest):")
            det_categories = [
                "Positive", "Negative", "Uncertainty", "Litigious",
                "Strong Modal", "Weak Modal", "Constraining", "Complexity"
            ]
            det_rows = []
            for r in analysis_data:
                row_det = {
                    "Fiscal Year": f"FY {r['year']}",
                    "Total Words": f"{r['total_words']:,d}" if r.get("total_words") is not None else "N/A",
                    "Positive": f"{r['positive']:,d}" if r.get("positive") is not None else "N/A",
                    "Negative": f"{r['negative']:,d}" if r.get("negative") is not None else "N/A"
                }
                for cat in det_categories[2:]:
                    val = r.get("other_sentiment", {}).get(cat)
                    row_det[cat] = f"{val:,d}" if val is not None else "N/A"
                row_det["P/N Ratio"] = f"{r['pn_ratio']:.2f}" if r.get("pn_ratio") is not None else "N/A"
                det_rows.append(row_det)
            st.dataframe(pd.DataFrame(det_rows), width="stretch", hide_index=True)

        # 5. OTHER SENTIMENT CATEGORY TRENDS
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("📊 Additional MD&A Sentiment Trends")
        st.caption("Select domain-specific sentiment categories to examine their trajectory across fiscal reporting periods:")

        multiselect_options = ["Uncertainty", "Litigious", "Strong Modal", "Weak Modal", "Constraining", "Complexity"]
        avail_options = [c for c in multiselect_options if c in available_cats]
        default_selections = [c for c in ["Uncertainty", "Litigious"] if c in avail_options]

        selected_add_cats = st.multiselect(
            "Select sentiment categories to plot:",
            options=avail_options,
            default=default_selections,
            key=f"{prefix}_{ticker}_add_sent_select"
        )

        if selected_add_cats:
            fig_add = go.Figure()
            palette = (
                ["#4C8DFF", "#F2C94C", "#B794F4", "#FF7A59", "#27C499", "#FF5C5C"]
                if is_dark else
                ["#4A5568", "#D69E2E", "#805AD5", "#DD6B20", "#319795", "#E53E3E"]
            )
            for idx, cat in enumerate(selected_add_cats):
                cat_series = [r.get("other_sentiment", {}).get(cat, 0) for r in analysis_data]
                fig_add.add_trace(go.Scatter(
                    x=chrono_years,
                    y=cat_series,
                    mode="lines+markers",
                    name=cat,
                    line=dict(color=palette[idx % len(palette)], width=2.5),
                    marker=dict(size=7),
                    hovertemplate=f"<b>%{{x}}</b><br>{cat}: %{{y:,d}} words<extra></extra>"
                ))
            apply_chart_theme(
                fig_add,
                title=f"Additional MD&A Sentiment Trends — {company_name} ({ticker})",
                xaxis_title="Fiscal Reporting Period",
                yaxis_title="Word Count (Item 7 MD&A)",
                height=380,
                margin=dict(l=60, r=40, t=60, b=50)
            )
            st.plotly_chart(fig_add, width="stretch", key=f"{prefix}_{ticker}_add_sent_chart")

        dict_word_count = results.get("dict_info", {}).get("total_words", 0)
        st.caption(f"Dictionary Indexed: Loughran-McDonald Master Sentiment Dictionary ({dict_word_count:,d} words).")

    # --------------------------------------------------------------------------
    # TAB 6: GROWTH & PERFORMANCE
    # --------------------------------------------------------------------------
    with tab_growth:
        st.subheader("📊 Long-Term Revenue & Earnings Growth")
        by_year = raw_accounts_info.get("by_year", {})
        stmt_unit = statement_data.get("Income Statement", {}).get("unit", "")
        if stmt_unit:
            st.caption(f"**Financial Statement Scale / Currency:** *{stmt_unit}*")

        growth_years = []
        rev_vals = []
        ni_vals = []
        for r in analysis_data:
            yr = r["year"]
            item = by_year.get(yr, {})
            rev = item.get("revenue")
            ni = item.get("net_income")
            if rev is not None and ni is not None:
                growth_years.append(f"FY {yr}")
                rev_vals.append(rev)
                ni_vals.append(ni)

        col_g1, col_g2 = st.columns(2)

        # GRAPH 1: Revenue vs Net Income
        with col_g1:
            fig_rev_ni = go.Figure()
            fig_rev_ni.add_trace(go.Bar(
                x=growth_years, y=rev_vals, name="Revenue", marker_color=THEME["chart_primary_line"],
                hovertemplate="<b>%{x}</b><br>Revenue: $%{y:,.0f}<extra></extra>"
            ))
            fig_rev_ni.add_trace(go.Bar(
                x=growth_years, y=ni_vals, name="Net Income", marker_color=THEME["chart_positive_color"],
                hovertemplate="<b>%{x}</b><br>Net Income: $%{y:,.0f}<extra></extra>"
            ))
            apply_chart_theme(
                fig_rev_ni,
                title="1. Revenue vs. Net Income Comparison",
                xaxis_title="Fiscal Year",
                yaxis_title=f"Amount ({stmt_unit})" if stmt_unit else "Amount ($)",
                barmode="group",
                height=350,
                margin=dict(l=50, r=20, t=60, b=40)
            )
            st.plotly_chart(fig_rev_ni, width="stretch", key=f"{prefix}_{ticker}_rev_ni")

        # GRAPH 2: Year-over-Year Revenue Growth Trend (%)
        with col_g2:
            rev_yoy_years = []
            rev_yoy_vals = []
            for i in range(1, len(rev_vals)):
                r_p = rev_vals[i-1]
                r_c = rev_vals[i]
                if r_p and r_p > 0:
                    yoy = ((r_c / r_p) - 1.0) * 100.0
                    rev_yoy_years.append(growth_years[i])
                    rev_yoy_vals.append(yoy)

            if rev_yoy_years:
                yoy_colors = [THEME["chart_positive_color"] if v >= 0 else THEME["chart_negative_color"] for v in rev_yoy_vals]
                fig_rev_growth = go.Figure(go.Bar(
                    x=rev_yoy_years, y=rev_yoy_vals, marker_color=yoy_colors,
                    hovertemplate="<b>%{x}</b><br>YoY Revenue Growth: %{y:+.2f}%<extra></extra>"
                ))
                apply_chart_theme(
                    fig_rev_growth,
                    title="2. YoY Revenue Growth Trend (%)",
                    xaxis_title="Fiscal Year",
                    yaxis_title="Revenue Growth (%)",
                    height=350,
                    show_legend=False,
                    margin=dict(l=50, r=20, t=60, b=40),
                    ticksuffix="%"
                )
                st.plotly_chart(fig_rev_growth, width="stretch", key=f"{prefix}_{ticker}_rev_growth")

        # GRAPH 3: Net Income Trend
        st.markdown("<br>", unsafe_allow_html=True)
        fig_ni_trend = go.Figure(go.Scatter(
            x=growth_years, y=ni_vals, mode="lines+markers",
            line=dict(color=THEME["chart_positive_color"], width=3), marker=dict(size=8, symbol="diamond"),
            hovertemplate="<b>%{x}</b><br>Net Income: $%{y:,.0f}<extra></extra>"
        ))
        apply_chart_theme(
            fig_ni_trend,
            title=f"3. Net Income Trajectory ({stmt_unit})" if stmt_unit else "3. Net Income Trajectory",
            xaxis_title="Fiscal Year",
            yaxis_title=f"Net Income ({stmt_unit})" if stmt_unit else "Net Income ($)",
            height=320,
            show_legend=False,
            margin=dict(l=50, r=20, t=50, b=40)
        )
        st.plotly_chart(fig_ni_trend, width="stretch", key=f"{prefix}_{ticker}_ni_trend")

        # Operating Summary Table
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("📋 Operating Performance Summary")
        growth_table_rows = []
        for i, yr_label in enumerate(growth_years):
            rev_c = rev_vals[i]
            ni_c = ni_vals[i]
            if i > 0 and rev_vals[i-1] > 0:
                yoy_rev_str = f"{((rev_c / rev_vals[i-1]) - 1.0) * 100.0:+.2f}%"
            else:
                yoy_rev_str = "Baseline"
            if i > 0 and ni_vals[i-1] != 0:
                yoy_ni_str = f"{((ni_c / abs(ni_vals[i-1])) - 1.0) * 100.0:+.2f}%"
            else:
                yoy_ni_str = "Baseline"

            growth_table_rows.append({
                "Fiscal Year": yr_label,
                "Revenue": f"${rev_c:,.0f}",
                "YoY Revenue Growth": yoy_rev_str,
                "Net Income": f"${ni_c:,.0f}",
                "YoY Net Income Growth": yoy_ni_str
            })
        st.dataframe(pd.DataFrame(growth_table_rows), width="stretch", hide_index=True)

    # --------------------------------------------------------------------------
    # TAB 7: DOWNLOAD REPORT
    # --------------------------------------------------------------------------
    with tab_download:
        st.subheader("📥 Download Master Financial Analysis Workbook")
        st.markdown(f"""
        Download the verified multi-worksheet Excel financial analysis workbook for **{company_name}**:
        - **Workbook File Name:** `{Path(output_filepath).name}`
        - **Worksheets Included:**
          1. `Income Statement` (Multi-year operational results)
          2. `Balance Sheet` (Multi-year assets, liabilities, and shareholders' equity)
          3. `Cash Flows` (Operating, investing, and financing cash flows)
          4. `Sentiment Analysis` (Loughran-McDonald category counts and P/N ratios)
          5. `Stock Prices` (Historical split-adjusted closing prices matched to balance sheet dates)
          6. `Analysis` (Chronological financial ratios, formula documentation, and dual-axis chart)
        """)

        if Path(output_filepath).exists():
            with open(output_filepath, "rb") as f:
                excel_bytes = f.read()

            st.download_button(
                label=f"📥 Download {Path(output_filepath).name}",
                data=excel_bytes,
                file_name=Path(output_filepath).name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width="stretch",
                key=f"{prefix}_{ticker}_dl_btn"
            )
            st.success("Master workbook is ready for immediate download.")
        else:
            st.error(f"Workbook file '{output_filepath}' not found on disk.")


# ==============================================================================
# RENDER APPLE VS. TESLA COMPARISON DASHBOARD
# ==============================================================================

def render_comparison_dashboard(aapl_results, tsla_results, target_filing_count):
    aapl_data = aapl_results["analysis_data"]
    tsla_data = tsla_results["analysis_data"]
    aapl_by_year = aapl_results["raw_accounts_info"].get("by_year", {})
    tsla_by_year = tsla_results["raw_accounts_info"].get("by_year", {})

    aapl_latest = aapl_data[-1]
    tsla_latest = tsla_data[-1]

    # Find common / aligned fiscal years
    aapl_years = [r["year"] for r in aapl_data]
    tsla_years = [r["year"] for r in tsla_data]
    common_years = sorted(list(set(aapl_years).intersection(tsla_years)))
    if not common_years:
        common_years = sorted(list(set(aapl_years).union(tsla_years)))

    aapl_dict = {r["year"]: r for r in aapl_data}
    tsla_dict = {r["year"]: r for r in tsla_data}

    # Comparison Header
    st.markdown(f"""
    <div class="company-header-card">
        <div style="display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 0.85rem;">
            <div>
                <div style="font-size: 1.7rem; font-weight: 700; color: {THEME['text_primary']}; line-height: 1.25;">
                    🍎 Apple Inc. (AAPL) <span style="font-size: 1.2rem; color: {THEME['text_muted']}; font-weight: 400;">vs.</span> ⚡ Tesla, Inc. (TSLA)
                </div>
                <div style="font-size: 0.92rem; color: {THEME['text_secondary']}; margin-top: 0.35rem; line-height: 1.4;">
                    Comparative Executive Financial & Sentiment Dashboard &nbsp;•&nbsp; Horizon: <b>{target_filing_count} Comparable Fiscal Years</b>
                </div>
            </div>
            <div style="display: flex; gap: 0.5rem; align-items: center; flex-wrap: wrap;">
                <span class="company-badge">🍎 AAPL: CIK {str(aapl_results['cik']).zfill(10)}</span>
                <span class="company-badge-secondary">⚡ TSLA: CIK {str(tsla_results['cik']).zfill(10)}</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Side-by-Side Latest-Year KPI Cards
    st.subheader("📌 Latest Fiscal-Year Side-by-Side Comparison")
    st.caption("Neutral comparative summary of latest reported accounting health, market price, and MD&A sentiment metrics.")

    k_c1, k_c2, k_c3 = st.columns(3)
    k_c4, k_c5, k_c6 = st.columns(3)

    # 1. Stock Price
    aapl_p = f"${aapl_latest['stock_price']:,.2f}" if aapl_latest.get("stock_price") is not None else "N/A"
    tsla_p = f"${tsla_latest['stock_price']:,.2f}" if tsla_latest.get("stock_price") is not None else "N/A"
    k_c1.markdown(f"""
    <div class="metric-card">
        <div class="metric-category">Market Value</div>
        <div class="metric-name">Stock Price</div>
        <div style="display:flex; justify-content:space-around; align-items:baseline; margin-top:0.4rem;">
            <div><span style="font-size:0.82rem; color:{THEME['aapl_color']}; font-weight:600;">AAPL:</span> <b style="font-size:1.25rem; color:{THEME['aapl_color']};">{aapl_p}</b></div>
            <div><span style="font-size:0.82rem; color:{THEME['tsla_color']}; font-weight:600;">TSLA:</span> <b style="font-size:1.25rem; color:{THEME['tsla_color']};">{tsla_p}</b></div>
        </div>
        <div class="metric-sub">FY {aapl_latest['year']} / FY {tsla_latest['year']} Balance Sheet Date</div>
    </div>
    """, unsafe_allow_html=True)

    # 2. Current Ratio
    aapl_cr = f"{aapl_latest['current_ratio']:.2f}" if aapl_latest.get("current_ratio") is not None else "N/A"
    tsla_cr = f"{tsla_latest['current_ratio']:.2f}" if tsla_latest.get("current_ratio") is not None else "N/A"
    k_c2.markdown(f"""
    <div class="metric-card">
        <div class="metric-category">Liquidity</div>
        <div class="metric-name">Current Ratio</div>
        <div style="display:flex; justify-content:space-around; align-items:baseline; margin-top:0.4rem;">
            <div><span style="font-size:0.82rem; color:{THEME['aapl_color']}; font-weight:600;">AAPL:</span> <b style="font-size:1.25rem; color:{THEME['aapl_color']};">{aapl_cr}</b></div>
            <div><span style="font-size:0.82rem; color:{THEME['tsla_color']}; font-weight:600;">TSLA:</span> <b style="font-size:1.25rem; color:{THEME['tsla_color']};">{tsla_cr}</b></div>
        </div>
        <div class="metric-sub">Current Assets / Current Liabilities</div>
    </div>
    """, unsafe_allow_html=True)

    # 3. Debt-to-Assets
    aapl_da = f"{aapl_latest['debt_to_assets'] * 100:.2f}%" if aapl_latest.get("debt_to_assets") is not None else "N/A"
    tsla_da = f"{tsla_latest['debt_to_assets'] * 100:.2f}%" if tsla_latest.get("debt_to_assets") is not None else "N/A"
    k_c3.markdown(f"""
    <div class="metric-card">
        <div class="metric-category">Solvency</div>
        <div class="metric-name">Debt-to-Assets</div>
        <div style="display:flex; justify-content:space-around; align-items:baseline; margin-top:0.4rem;">
            <div><span style="font-size:0.82rem; color:{THEME['aapl_color']}; font-weight:600;">AAPL:</span> <b style="font-size:1.25rem; color:{THEME['aapl_color']};">{aapl_da}</b></div>
            <div><span style="font-size:0.82rem; color:{THEME['tsla_color']}; font-weight:600;">TSLA:</span> <b style="font-size:1.25rem; color:{THEME['tsla_color']};">{tsla_da}</b></div>
        </div>
        <div class="metric-sub">Total Liabilities / Total Assets</div>
    </div>
    """, unsafe_allow_html=True)

    # 4. Total Asset Turnover
    aapl_at = f"{aapl_latest['asset_turnover']:.2f}" if aapl_latest.get("asset_turnover") is not None else "N/A"
    tsla_at = f"{tsla_latest['asset_turnover']:.2f}" if tsla_latest.get("asset_turnover") is not None else "N/A"
    k_c4.markdown(f"""
    <div class="metric-card">
        <div class="metric-category">Efficiency</div>
        <div class="metric-name">Asset Turnover</div>
        <div style="display:flex; justify-content:space-around; align-items:baseline; margin-top:0.4rem;">
            <div><span style="font-size:0.82rem; color:{THEME['aapl_color']}; font-weight:600;">AAPL:</span> <b style="font-size:1.25rem; color:{THEME['aapl_color']};">{aapl_at}</b></div>
            <div><span style="font-size:0.82rem; color:{THEME['tsla_color']}; font-weight:600;">TSLA:</span> <b style="font-size:1.25rem; color:{THEME['tsla_color']};">{tsla_at}</b></div>
        </div>
        <div class="metric-sub">Revenue / Average Assets</div>
    </div>
    """, unsafe_allow_html=True)

    # 5. Net Profit Margin
    aapl_pm = f"{aapl_latest['profit_margin'] * 100:.2f}%" if aapl_latest.get("profit_margin") is not None else "N/A"
    tsla_pm = f"{tsla_latest['profit_margin'] * 100:.2f}%" if tsla_latest.get("profit_margin") is not None else "N/A"
    k_c5.markdown(f"""
    <div class="metric-card">
        <div class="metric-category">Profitability</div>
        <div class="metric-name">Net Margin</div>
        <div style="display:flex; justify-content:space-around; align-items:baseline; margin-top:0.4rem;">
            <div><span style="font-size:0.82rem; color:{THEME['aapl_color']}; font-weight:600;">AAPL:</span> <b style="font-size:1.25rem; color:{THEME['aapl_color']};">{aapl_pm}</b></div>
            <div><span style="font-size:0.82rem; color:{THEME['tsla_color']}; font-weight:600;">TSLA:</span> <b style="font-size:1.25rem; color:{THEME['tsla_color']};">{tsla_pm}</b></div>
        </div>
        <div class="metric-sub">Net Income / Revenue</div>
    </div>
    """, unsafe_allow_html=True)

    # 6. MD&A P/N Ratio
    aapl_pn = f"{aapl_latest['pn_ratio']:.2f}" if aapl_latest.get("pn_ratio") is not None else "N/A"
    tsla_pn = f"{tsla_latest['pn_ratio']:.2f}" if tsla_latest.get("pn_ratio") is not None else "N/A"
    k_c6.markdown(f"""
    <div class="metric-card">
        <div class="metric-category">MD&A Sentiment</div>
        <div class="metric-name">P/N Ratio</div>
        <div style="display:flex; justify-content:space-around; align-items:baseline; margin-top:0.4rem;">
            <div><span style="font-size:0.82rem; color:{THEME['aapl_color']}; font-weight:600;">AAPL:</span> <b style="font-size:1.25rem; color:{THEME['aapl_color']};">{aapl_pn}</b></div>
            <div><span style="font-size:0.82rem; color:{THEME['tsla_color']}; font-weight:600;">TSLA:</span> <b style="font-size:1.25rem; color:{THEME['tsla_color']};">{tsla_pn}</b></div>
        </div>
        <div class="metric-sub">Positive Words / Negative Words</div>
    </div>
    """, unsafe_allow_html=True)

    # COMPARISON GRAPH 1: Normalized Stock Performance
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("1. Indexed Stock Performance — Base Year = 100")
    st.caption("Normalized historical stock price trajectory indexing the initial comparative fiscal year to 100. Eliminates nominal share price differences.")

    chart_years = [str(yr) for yr in common_years]
    aapl_p0 = aapl_dict.get(common_years[0], {}).get("stock_price")
    tsla_p0 = tsla_dict.get(common_years[0], {}).get("stock_price")

    aapl_stock_idx = []
    tsla_stock_idx = []
    for yr in common_years:
        ap = aapl_dict.get(yr, {}).get("stock_price")
        tp = tsla_dict.get(yr, {}).get("stock_price")
        aapl_stock_idx.append(((ap / aapl_p0) * 100.0) if ap and aapl_p0 and aapl_p0 > 0 else None)
        tsla_stock_idx.append(((tp / tsla_p0) * 100.0) if tp and tsla_p0 and tsla_p0 > 0 else None)

    fig_idx_stock = go.Figure()
    fig_idx_stock.add_trace(go.Scatter(
        x=chart_years, y=aapl_stock_idx, mode="lines+markers",
        name="Apple (AAPL)", line=dict(color=THEME["aapl_color"], width=3), marker=dict(size=8, symbol="circle"),
        hovertemplate="<b>AAPL (FY %{x})</b><br>Index: %{y:.1f}<extra></extra>"
    ))
    fig_idx_stock.add_trace(go.Scatter(
        x=chart_years, y=tsla_stock_idx, mode="lines+markers",
        name="Tesla (TSLA)", line=dict(color=THEME["tsla_color"], width=3), marker=dict(size=8, symbol="diamond"),
        hovertemplate="<b>TSLA (FY %{x})</b><br>Index: %{y:.1f}<extra></extra>"
    ))
    fig_idx_stock.add_hline(
        y=100.0, line_dash="dash", line_color=THEME["chart_neutral_line"],
        annotation_text="Base Year = 100", annotation_font_color=THEME["text_secondary"]
    )
    apply_chart_theme(
        fig_idx_stock,
        title="Indexed Stock Performance — Base Year = 100",
        xaxis_title="Fiscal Year",
        yaxis_title="Indexed Performance (Base=100)",
        height=400,
        margin=dict(l=50, r=20, t=60, b=40)
    )
    st.plotly_chart(fig_idx_stock, width="stretch", key="comp_chart_idx_stock")

    # COMPARISON GRAPH 2: MD&A Sentiment Comparison
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("2. MD&A Sentiment P/N Ratio Comparison")
    st.caption("Comparative tone analysis of Item 7 MD&A filings across identical fiscal years.")

    aapl_pn_series = [aapl_dict.get(yr, {}).get("pn_ratio") for yr in common_years]
    tsla_pn_series = [tsla_dict.get(yr, {}).get("pn_ratio") for yr in common_years]

    fig_pn_comp = go.Figure()
    fig_pn_comp.add_trace(go.Scatter(
        x=chart_years, y=aapl_pn_series, mode="lines+markers",
        name="Apple P/N Ratio", line=dict(color=THEME["aapl_color"], width=2.5), marker=dict(size=8, symbol="circle"),
        hovertemplate="<b>AAPL (FY %{x})</b><br>P/N Ratio: %{y:.2f}<extra></extra>"
    ))
    fig_pn_comp.add_trace(go.Scatter(
        x=chart_years, y=tsla_pn_series, mode="lines+markers",
        name="Tesla P/N Ratio", line=dict(color=THEME["tsla_color"], width=2.5), marker=dict(size=8, symbol="diamond"),
        hovertemplate="<b>TSLA (FY %{x})</b><br>P/N Ratio: %{y:.2f}<extra></extra>"
    ))
    fig_pn_comp.add_hline(
        y=1.0, line_dash="dash", line_color=THEME["chart_neutral_line"],
        annotation_text="P/N = 1.0 (Neutral Baseline)", annotation_font_color=THEME["text_secondary"]
    )
    apply_chart_theme(
        fig_pn_comp,
        title="MD&A Sentiment P/N Ratio Comparison",
        xaxis_title="Fiscal Year",
        yaxis_title="P/N Ratio",
        height=380,
        margin=dict(l=50, r=20, t=60, b=40),
        tickformat=".2f"
    )
    st.plotly_chart(fig_pn_comp, width="stretch", key="comp_chart_pn_line")

    # COMPARISON GRAPH 3: Financial Ratio Comparison
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("3. Financial Ratio Comparison")
    st.caption("Select a key financial ratio dimension to view Apple and Tesla trajectories on the same chart:")

    ratio_choice = st.selectbox(
        "Select Ratio Dimension:",
        options=[
            "Current Ratio (Liquidity)",
            "Debt-to-Assets Ratio (Solvency)",
            "Total Asset Turnover (Efficiency)",
            "Net Profit Margin (Profitability)"
        ],
        key="comp_ratio_selector"
    )

    if ratio_choice == "Current Ratio (Liquidity)":
        aapl_r_vals = [aapl_dict.get(yr, {}).get("current_ratio") for yr in common_years]
        tsla_r_vals = [tsla_dict.get(yr, {}).get("current_ratio") for yr in common_years]
        y_title = "Current Ratio"
        y_fmt = ".2f"
        val_hover = "%{y:.2f}"
    elif ratio_choice == "Debt-to-Assets Ratio (Solvency)":
        aapl_r_vals = [(aapl_dict.get(yr, {}).get("debt_to_assets") * 100.0) if aapl_dict.get(yr, {}).get("debt_to_assets") is not None else None for yr in common_years]
        tsla_r_vals = [(tsla_dict.get(yr, {}).get("debt_to_assets") * 100.0) if tsla_dict.get(yr, {}).get("debt_to_assets") is not None else None for yr in common_years]
        y_title = "Debt-to-Assets (%)"
        y_fmt = ".1f"
        val_hover = "%{y:.2f}%"
    elif ratio_choice == "Total Asset Turnover (Efficiency)":
        aapl_r_vals = [aapl_dict.get(yr, {}).get("asset_turnover") for yr in common_years]
        tsla_r_vals = [tsla_dict.get(yr, {}).get("asset_turnover") for yr in common_years]
        y_title = "Asset Turnover"
        y_fmt = ".2f"
        val_hover = "%{y:.2f}"
    else:  # Net Profit Margin
        aapl_r_vals = [(aapl_dict.get(yr, {}).get("profit_margin") * 100.0) if aapl_dict.get(yr, {}).get("profit_margin") is not None else None for yr in common_years]
        tsla_r_vals = [(tsla_dict.get(yr, {}).get("profit_margin") * 100.0) if tsla_dict.get(yr, {}).get("profit_margin") is not None else None for yr in common_years]
        y_title = "Net Margin (%)"
        y_fmt = ".1f"
        val_hover = "%{y:.2f}%"

    fig_ratio_comp = go.Figure()
    fig_ratio_comp.add_trace(go.Scatter(
        x=chart_years, y=aapl_r_vals, mode="lines+markers",
        name="Apple (AAPL)", line=dict(color=THEME["aapl_color"], width=2.5), marker=dict(size=8, symbol="circle"),
        hovertemplate=f"<b>AAPL (FY %{{x}})</b><br>{ratio_choice}: {val_hover}<extra></extra>"
    ))
    fig_ratio_comp.add_trace(go.Scatter(
        x=chart_years, y=tsla_r_vals, mode="lines+markers",
        name="Tesla (TSLA)", line=dict(color=THEME["tsla_color"], width=2.5), marker=dict(size=8, symbol="diamond"),
        hovertemplate=f"<b>TSLA (FY %{{x}})</b><br>{ratio_choice}: {val_hover}<extra></extra>"
    ))
    apply_chart_theme(
        fig_ratio_comp,
        title=f"{ratio_choice} — Apple vs. Tesla",
        xaxis_title="Fiscal Year",
        yaxis_title=y_title,
        height=380,
        margin=dict(l=50, r=20, t=60, b=40)
    )
    st.plotly_chart(fig_ratio_comp, width="stretch", key="comp_chart_ratio_line")

    # COMPARISON GRAPH 4: Revenue Growth Index
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("4. Indexed Revenue Growth — Base Year = 100")
    st.caption("Normalized top-line revenue expansion indexing the initial comparative period to 100. Enables clear growth rate comparison regardless of scale.")

    aapl_r0 = aapl_by_year.get(common_years[0], {}).get("revenue")
    tsla_r0 = tsla_by_year.get(common_years[0], {}).get("revenue")

    aapl_rev_idx = []
    tsla_rev_idx = []
    for yr in common_years:
        ar = aapl_by_year.get(yr, {}).get("revenue")
        tr = tsla_by_year.get(yr, {}).get("revenue")
        aapl_rev_idx.append(((ar / aapl_r0) * 100.0) if ar and aapl_r0 and aapl_r0 > 0 else None)
        tsla_rev_idx.append(((tr / tsla_r0) * 100.0) if tr and tsla_r0 and tsla_r0 > 0 else None)

    fig_idx_rev = go.Figure()
    fig_idx_rev.add_trace(go.Scatter(
        x=chart_years, y=aapl_rev_idx, mode="lines+markers",
        name="Apple Revenue Index", line=dict(color=THEME["aapl_color"], width=3), marker=dict(size=8, symbol="circle"),
        hovertemplate="<b>AAPL Revenue Index (FY %{x})</b><br>Index: %{y:.1f}<extra></extra>"
    ))
    fig_idx_rev.add_trace(go.Scatter(
        x=chart_years, y=tsla_rev_idx, mode="lines+markers",
        name="Tesla Revenue Index", line=dict(color=THEME["tsla_color"], width=3), marker=dict(size=8, symbol="diamond"),
        hovertemplate="<b>TSLA Revenue Index (FY %{x})</b><br>Index: %{y:.1f}<extra></extra>"
    ))
    fig_idx_rev.add_hline(
        y=100.0, line_dash="dash", line_color=THEME["chart_neutral_line"],
        annotation_text="Base Year = 100", annotation_font_color=THEME["text_secondary"]
    )
    apply_chart_theme(
        fig_idx_rev,
        title="Indexed Revenue Growth — Base Year = 100",
        xaxis_title="Fiscal Year",
        yaxis_title="Indexed Revenue (Base=100)",
        height=400,
        margin=dict(l=50, r=20, t=60, b=40)
    )
    st.plotly_chart(fig_idx_rev, width="stretch", key="comp_chart_idx_rev")

    # CROSS-COMPANY SUMMARY TABLE
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("📋 Cross-Company Chronological Summary Table")
    st.caption("Comprehensive chronological table comparing Apple and Tesla across stock prices, ratios, and sentiment scores (oldest to newest):")

    comp_rows = []
    for yr in common_years:
        ar = aapl_dict.get(yr)
        tr = tsla_dict.get(yr)
        comp_rows.append({
            "Fiscal Year": f"FY {yr}",
            "AAPL Stock Price": f"${ar['stock_price']:,.2f}" if ar and ar.get("stock_price") is not None else "N/A",
            "TSLA Stock Price": f"${tr['stock_price']:,.2f}" if tr and tr.get("stock_price") is not None else "N/A",
            "AAPL Current Ratio": f"{ar['current_ratio']:.2f}" if ar and ar.get("current_ratio") is not None else "N/A",
            "TSLA Current Ratio": f"{tr['current_ratio']:.2f}" if tr and tr.get("current_ratio") is not None else "N/A",
            "AAPL Debt-to-Assets": f"{ar['debt_to_assets'] * 100:.2f}%" if ar and ar.get("debt_to_assets") is not None else "N/A",
            "TSLA Debt-to-Assets": f"{tr['debt_to_assets'] * 100:.2f}%" if tr and tr.get("debt_to_assets") is not None else "N/A",
            "AAPL Asset Turnover": f"{ar['asset_turnover']:.2f}" if ar and ar.get("asset_turnover") is not None else "N/A",
            "TSLA Asset Turnover": f"{tr['asset_turnover']:.2f}" if tr and tr.get("asset_turnover") is not None else "N/A",
            "AAPL Net Profit Margin": f"{ar['profit_margin'] * 100:.2f}%" if ar and ar.get("profit_margin") is not None else "N/A",
            "TSLA Net Profit Margin": f"{tr['profit_margin'] * 100:.2f}%" if tr and tr.get("profit_margin") is not None else "N/A",
            "AAPL P/N Ratio": f"{ar['pn_ratio']:.2f}" if ar and ar.get("pn_ratio") is not None else "N/A",
            "TSLA P/N Ratio": f"{tr['pn_ratio']:.2f}" if tr and tr.get("pn_ratio") is not None else "N/A",
        })
    st.dataframe(pd.DataFrame(comp_rows), width="stretch", hide_index=True)

    # WORKBOOK DOWNLOADS
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("📥 Download Analysis Workbooks")
    st.caption("Download the individual verified Excel workbooks generated for each company:")

    dl_c1, dl_c2 = st.columns(2)
    aapl_wb = aapl_results.get("output_filepath")
    tsla_wb = tsla_results.get("output_filepath")

    with dl_c1:
        if aapl_wb and Path(aapl_wb).exists():
            with open(aapl_wb, "rb") as f:
                b_aapl = f.read()
            st.download_button(
                label=f"📥 Download Apple Workbook ({Path(aapl_wb).name})",
                data=b_aapl,
                file_name=Path(aapl_wb).name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width="stretch",
                key="dl_comp_aapl"
            )

    with dl_c2:
        if tsla_wb and Path(tsla_wb).exists():
            with open(tsla_wb, "rb") as f:
                b_tsla = f.read()
            st.download_button(
                label=f"📥 Download Tesla Workbook ({Path(tsla_wb).name})",
                data=b_tsla,
                file_name=Path(tsla_wb).name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width="stretch",
                key="dl_comp_tsla"
            )


# ==============================================================================
# MAIN VIEWING AREA: SWITCHABLE TOP-LEVEL COMPANY DASHBOARDS
# ==============================================================================

target_filing_count = int(st.session_state.get("selected_filing_count", 4))
custom_ticker_current = st.session_state.get("custom_ticker", "AAPL").strip().upper()

top_tab_aapl, top_tab_tsla, top_tab_comp, top_tab_custom = st.tabs([
    "🍎 Apple (AAPL)",
    "⚡ Tesla (TSLA)",
    "📊 Apple vs. Tesla",
    "🔎 Custom Company"
])

# ------------------------------------------------------------------------------
# TOP TAB 1: APPLE (AAPL)
# ------------------------------------------------------------------------------
with top_tab_aapl:
    aapl_cached = get_cached_result("AAPL", target_filing_count)
    if aapl_cached:
        render_company_dashboard(aapl_cached, target_filing_count, prefix="aapl")
    else:
        st.info(f"💡 **Apple Inc. (AAPL)** has not been analyzed for this **{target_filing_count}-year horizon** yet. Select Apple in the sidebar Analysis Settings and click **🚀 Run Analysis** to generate the report.")

# ------------------------------------------------------------------------------
# TOP TAB 2: TESLA (TSLA)
# ------------------------------------------------------------------------------
with top_tab_tsla:
    tsla_cached = get_cached_result("TSLA", target_filing_count)
    if tsla_cached:
        render_company_dashboard(tsla_cached, target_filing_count, prefix="tsla")
    else:
        st.info(f"💡 **Tesla, Inc. (TSLA)** has not been analyzed for this **{target_filing_count}-year horizon** yet. Select Tesla in the sidebar Analysis Settings and click **🚀 Run Analysis** to generate the report.")

# ------------------------------------------------------------------------------
# TOP TAB 3: APPLE VS. TESLA COMPARISON
# ------------------------------------------------------------------------------
with top_tab_comp:
    aapl_comp = get_cached_result("AAPL", target_filing_count)
    tsla_comp = get_cached_result("TSLA", target_filing_count)

    if aapl_comp and tsla_comp:
        render_comparison_dashboard(aapl_comp, tsla_comp, target_filing_count)
    else:
        st.info(f"💡 **Apple vs. Tesla Comparison requires both companies to be analyzed for the same {target_filing_count}-year horizon.**")
        st.markdown(f"""
        **Current Status for {target_filing_count}-Year Analysis Horizon:**
        - 🍎 **Apple Inc. (AAPL):** {'✅ Ready for comparison' if aapl_comp else f'❌ Not analyzed yet for {target_filing_count} filings'}
        - ⚡ **Tesla, Inc. (TSLA):** {'✅ Ready for comparison' if tsla_comp else f'❌ Not analyzed yet for {target_filing_count} filings'}

        *Select the missing company in the sidebar Analysis Settings and click **🚀 Run Analysis** to generate its report. Once both are complete, this comparison will instantly activate.*
        """)

# ------------------------------------------------------------------------------
# TOP TAB 4: CUSTOM COMPANY
# ------------------------------------------------------------------------------
with top_tab_custom:
    if custom_ticker_current in ("AAPL", "TSLA"):
        st.info("💡 This tab displays analyses for custom tickers (e.g. MSFT, NVDA, AMZN). Apple and Tesla have dedicated tabs above. Select **Custom Ticker** in the sidebar settings and enter another company's symbol to generate its report.")
    else:
        custom_cached = get_cached_result(custom_ticker_current, target_filing_count)
        if custom_cached:
            render_company_dashboard(custom_cached, target_filing_count, prefix="custom")
        else:
            st.info(f"💡 **Custom Company ({custom_ticker_current})** has not been analyzed for this **{target_filing_count}-year horizon** yet. Enter ticker symbol in sidebar Analysis Settings and click **🚀 Run Analysis** to generate the report.")
