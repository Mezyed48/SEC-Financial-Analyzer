"""
Streamlit Web Application: SEC 10-K Financial & Sentiment Analyzer

This local web frontend wraps the verified Steps 2-9 financial analysis pipeline,
providing an interactive interface for:
  - Form 10-K retrieval from SEC EDGAR
  - Item 7 MD&A plain-text extraction and cleaning
  - Financial statements extraction (Income Statement, Balance Sheet, Cash Flows)
  - Loughran-McDonald sentiment analysis (8 sentiment categories, P/N ratio)
  - Historical split-adjusted stock prices (Yahoo Finance via yfinance)
  - Financial ratio analysis (Liquidity, Solvency, Efficiency, Profitability)
  - Interactive dual-axis Plotly visualization (Stock Price vs. P/N Ratio)
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
from main import (
    run_analysis_workflow,
    PROJECT_ROOT,
    setup_dictionaries_directory,
    find_dictionary_file,
    load_loughran_mcdonald_dictionary
)


# ==============================================================================
# PAGE CONFIGURATION & STYLES
# ==============================================================================

st.set_page_config(
    page_title="SEC 10-K Financial & Sentiment Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom styling for KPI cards and headers
st.markdown("""
<style>
    .main-title {
        font-size: 2.1rem;
        font-weight: 700;
        color: #1F497D;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 1.1rem;
        font-weight: 400;
        color: #4A5568;
        margin-bottom: 0.8rem;
    }
    .intro-box {
        background-color: #F7FAFC;
        border-left: 4px solid #1F497D;
        padding: 0.9rem 1.2rem;
        border-radius: 6px;
        margin-bottom: 1.5rem;
        font-size: 0.95rem;
        color: #2D3748;
        line-height: 1.5;
    }
    .company-header-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1.25rem 1.5rem;
        margin-bottom: 1.25rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.03);
    }
    .company-badge {
        background-color: #EBF8FF;
        color: #2B6CB0;
        border: 1px solid #BEE3F8;
        padding: 0.35rem 0.85rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        display: inline-block;
    }
    .company-badge-secondary {
        background-color: #EDF2F7;
        color: #4A5568;
        border: 1px solid #E2E8F0;
        padding: 0.35rem 0.85rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        display: inline-block;
    }
    .metric-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 0.9rem 0.75rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        text-align: center;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .metric-category {
        font-size: 0.72rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #718096;
        margin-bottom: 0.2rem;
    }
    .metric-name {
        font-size: 0.88rem;
        font-weight: 600;
        color: #2D3748;
        margin-bottom: 0.35rem;
    }
    .metric-val {
        font-size: 1.55rem;
        font-weight: 700;
        color: #1F497D;
        line-height: 1.2;
    }
    .metric-sub {
        font-size: 0.74rem;
        color: #718096;
        margin-top: 0.3rem;
        line-height: 1.2;
    }
    .metric-sub.positive {
        color: #2F855A;
        font-weight: 600;
    }
    .metric-sub.negative {
        color: #C53030;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)


# ==============================================================================
# HEADER & APPLICATION OVERVIEW
# ==============================================================================

st.markdown('<div class="main-title">SEC 10-K Financial & Sentiment Analyzer</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Automated SEC filing analysis, financial ratios, MD&A sentiment, and historical stock prices</div>', unsafe_allow_html=True)

st.markdown("""
<div class="intro-box">
This tool retrieves SEC Form 10-K filings, extracts financial statement and MD&A data,
performs Loughran-McDonald financial sentiment analysis, calculates financial ratios,
retrieves historical stock prices, and generates a downloadable Excel analysis workbook.
</div>
""", unsafe_allow_html=True)


# ==============================================================================
# FEATURED COMPANY QUICK SWITCHER (EXECUTIVE DASHBOARD)
# ==============================================================================

# Initialize session state for company selection & multi-company analysis cache
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

st.markdown("##### 🏢 Featured Company Quick Switcher")
company_options = ["Apple (AAPL)", "Tesla (TSLA)", "Custom Ticker"]

selected_mode = st.pills(
    "Select Featured Company:",
    options=company_options,
    default=st.session_state.company_mode,
    key="company_pill_selector",
    label_visibility="collapsed"
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


# ==============================================================================
# LOUGHRAN-MCDONALD DICTIONARY STATUS & DETECTION
# ==============================================================================

dict_dir = setup_dictionaries_directory()
dict_filepath, dict_filename = find_dictionary_file(dict_dir)
has_dictionary = dict_filepath is not None and os.path.isfile(dict_filepath)


# ==============================================================================
# SIDEBAR / USER INPUT CONTROLS
# ==============================================================================

with st.sidebar:
    st.header("⚙️ Analysis Settings")
    st.markdown("Configure target public company and historical period horizon:")

    # Ensure sidebar ticker widget state strictly reflects active selection
    if st.session_state.company_mode == "Apple (AAPL)":
        st.session_state.sidebar_ticker_input = "AAPL"
        st.session_state.active_ticker = "AAPL"
        is_custom = False
        ticker_help = "Selected via top quick-switcher (Apple Inc. — AAPL)."
    elif st.session_state.company_mode == "Tesla (TSLA)":
        st.session_state.sidebar_ticker_input = "TSLA"
        st.session_state.active_ticker = "TSLA"
        is_custom = False
        ticker_help = "Selected via top quick-switcher (Tesla, Inc. — TSLA)."
    else:
        # Custom Ticker mode: restore custom ticker if previous value was a fixed featured ticker
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
        st.success(f"**Loaded:** `{dict_filename}`")
        with st.expander("Replace / Upload New Dictionary", expanded=False):
            st.caption("Upload an updated Loughran-McDonald dictionary (.csv / .xlsx):")
            uploaded_sidebar = st.file_uploader(
                "Upload replacement dictionary",
                type=["csv", "xlsx"],
                key="sidebar_dict_uploader",
                label_visibility="collapsed"
            )
            if uploaded_sidebar is not None:
                save_path = os.path.join(dict_dir, uploaded_sidebar.name)
                with open(save_path, "wb") as f_out:
                    f_out.write(uploaded_sidebar.getbuffer())
                check = load_loughran_mcdonald_dictionary(save_path)
                if check and check.get("total_words", 0) > 0:
                    st.success(f"Loaded {uploaded_sidebar.name} ({check['total_words']:,d} words)!")
                    st.rerun()
                else:
                    if os.path.exists(save_path):
                        os.remove(save_path)
                    st.error("Invalid dictionary file format. Please upload a .csv or .xlsx file.")
    else:
        st.warning("⚠️ **Dictionary Missing**\n\nUpload required for Step 7 sentiment analysis.")

    st.markdown("---")
    st.caption("🔒 **SEC Fair Access Compliant**: Requests are throttled and use declarative User-Agent contact credentials.")
    st.caption("📚 **Accounting Standards**: Standard financial-analysis ratios calculated using U.S. GAAP financial-statement data.")


# ==============================================================================
# DICTIONARY UPLOAD INTERFACE (FOR FRESH DEPLOYMENTS)
# ==============================================================================

if not has_dictionary:
    st.warning("### ⚠️ Loughran-McDonald Sentiment Dictionary Required")
    st.markdown("""
    To perform **Step 7 (MD&A Sentiment Analysis)**, this application requires the official **Loughran-McDonald Financial Sentiment Dictionary**.

    > **Notice on Copyright & Licensing:** Per academic guidelines, the proprietary dictionary file created by Tim Loughran and Bill McDonald is **not distributed** in the public GitHub repository.

    **How to obtain the official dictionary:**
    1. Visit the University of Notre Dame [Software Repository for Accounting and Finance (SRAF)](https://sraf.nd.edu/loughranmcdonald-master-dictionary/).
    2. Download the official **Loughran-McDonald Master Dictionary** (`.csv` or `.xlsx`).
    3. Upload the downloaded file below to activate financial sentiment analysis for your session.
    """)

    uploaded_main_dict = st.file_uploader(
        "Upload Loughran-McDonald Master Dictionary (.csv or .xlsx):",
        type=["csv", "xlsx"],
        key="main_dict_uploader",
        help="Upload the official Loughran-McDonald Master Dictionary obtained from Notre Dame SRAF."
    )

    if uploaded_main_dict is not None:
        save_path = os.path.join(dict_dir, uploaded_main_dict.name)
        try:
            with open(save_path, "wb") as f_out:
                f_out.write(uploaded_main_dict.getbuffer())

            # Validate the uploaded file using the existing dictionary-loading function
            check = load_loughran_mcdonald_dictionary(save_path)
            if check and check.get("total_words", 0) > 0:
                st.success(
                    f"✅ **Dictionary Loaded Successfully!** Indexed {check['total_words']:,d} words "
                    f"across {len(check['categories'])} sentiment categories ({', '.join(check['categories'])})."
                )
                st.rerun()
            else:
                if os.path.exists(save_path):
                    os.remove(save_path)
                st.error("❌ The uploaded file is not a valid Loughran-McDonald dictionary. Please ensure you upload the official CSV or Excel (.xlsx) file from Notre Dame SRAF.")
        except Exception as upload_err:
            if os.path.exists(save_path):
                os.remove(save_path)
            st.error(f"❌ Error saving uploaded dictionary: {upload_err}")


# ==============================================================================
# RUN ANALYSIS TRIGGER (RERUN-SAFE)
# ==============================================================================

if run_btn:
    # Input validation
    clean_ticker = st.session_state.active_ticker.strip().upper()
    if not clean_ticker or not re.match(r"^[A-Z0-9\.\-]+$", clean_ticker):
        st.error("❌ Please enter a valid stock ticker symbol (letters, numbers, dots, or hyphens only).")
    elif not has_dictionary:
        st.error("❌ Cannot start analysis: The Loughran-McDonald sentiment dictionary is required. Please upload the dictionary file above first.")
    else:
        # Progress tracking status container
        status_box = st.status(f"Starting SEC 10-K Analysis for **{clean_ticker}** ({input_filing_count} filings)...", expanded=True)
        progress_bar = st.progress(0.0)

        def on_progress(step_num, stage_name, description):
            pct = min(1.0, step_num / 10.0)
            progress_bar.progress(pct)
            status_box.write(f"**Step {step_num}/10: {stage_name}** — {description}")

        try:
            # Execute backend workflow without duplicating logic
            workflow_results = run_analysis_workflow(
                ticker=clean_ticker,
                filing_count=input_filing_count,
                progress_callback=on_progress,
                is_cli=False
            )

            if workflow_results.get("success"):
                progress_bar.progress(1.0)
                status_box.update(label=f"✅ Analysis for {clean_ticker} complete!", state="complete", expanded=False)
                # Store in multi-company cache by unique key (ticker, filing_count) and by ticker (Requirement 4)
                cache_key = (clean_ticker, input_filing_count)
                st.session_state.analysis_cache[cache_key] = workflow_results
                st.session_state.analysis_cache[clean_ticker] = workflow_results
            else:
                progress_bar.empty()
                status_box.update(label=f"❌ Analysis failed for {clean_ticker}", state="error", expanded=True)
                st.error(f"**Error:** {workflow_results.get('error')}")

        except Exception as ex:
            progress_bar.empty()
            status_box.update(label="❌ Unexpected Error", state="error", expanded=True)
            st.error(f"An unexpected error occurred during execution: {ex}")


# ==============================================================================
# DISPLAY RESULTS DASHBOARD
# ==============================================================================

active_ticker = st.session_state.get("active_ticker", "AAPL").strip().upper()
target_filing_count = int(st.session_state.get("selected_filing_count", 4))

# Lookup in cache: exact (ticker, filing_count) first, then fallback to most recent run for ticker
results = st.session_state.analysis_cache.get((active_ticker, target_filing_count))
if not results and active_ticker in st.session_state.analysis_cache:
    results = st.session_state.analysis_cache[active_ticker]

if results:
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

    # Polished Executive Company Header (Requirement 1 & 6)
    oldest_yr = analysis_data[0]["year"]
    latest_yr = analysis_data[-1]["year"]
    fiscal_horizon = f"FY {oldest_yr} – FY {latest_yr}"
    reporting_scope = f"{num_years}-Year SEC 10-K Analysis"
    cik_str = str(cik).zfill(10)

    st.markdown(f"""
    <div class="company-header-card">
        <div style="display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 0.85rem;">
            <div>
                <div style="font-size: 1.7rem; font-weight: 700; color: #1F497D; line-height: 1.25;">
                    🏢 {company_name} <span style="font-size: 1.2rem; color: #4A5568; font-weight: 500;">({ticker})</span>
                </div>
                <div style="font-size: 0.92rem; color: #4A5568; margin-top: 0.35rem; line-height: 1.4;">
                    SEC Central Index Key: <b>CIK {cik_str}</b> &nbsp;•&nbsp; Reporting Period: <b>{fiscal_horizon}</b> &nbsp;•&nbsp; Master Workbook: <b>{Path(output_filepath).name}</b>
                </div>
            </div>
            <div style="display: flex; gap: 0.5rem; align-items: center; flex-wrap: wrap;">
                <span class="company-badge">📋 {reporting_scope}</span>
                <span class="company-badge-secondary">📅 FY {oldest_yr} – FY {latest_yr}</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Six Navigation Tabs
    tab_overview, tab_ratios, tab_sentiment, tab_stock, tab_statements, tab_download = st.tabs([
        "📊 Overview",
        "📈 Financial Ratios",
        "📝 Sentiment Analysis",
        "💲 Stock Prices",
        "📑 Financial Statements",
        "📥 Download Report"
    ])

    # --------------------------------------------------------------------------
    # TAB 1: OVERVIEW & CHART
    # --------------------------------------------------------------------------
    with tab_overview:
        oldest = analysis_data[0]
        latest = analysis_data[-1]

        # Calculate Period Performance Metrics from existing workflow data (Requirement 2)
        # 1. Period Stock Return
        oldest_p = oldest.get("stock_price")
        latest_p = latest.get("stock_price")
        if oldest_p is not None and latest_p is not None and oldest_p > 0:
            stock_ret_pct = ((latest_p - oldest_p) / oldest_p) * 100.0
            stock_ret_str = f"{stock_ret_pct:+.1f}%"
            stock_ret_cls = "positive" if stock_ret_pct >= 0 else "negative"
            stock_ret_sub = f"FY {oldest['year']} (${oldest_p:,.2f}) → FY {latest['year']} (${latest_p:,.2f})"
        else:
            stock_ret_str = "N/A"
            stock_ret_cls = ""
            stock_ret_sub = "Price data unavailable"

        # 2. Period Revenue Growth (from raw_accounts_info / statement data)
        by_year = raw_accounts_info.get("by_year", {})
        oldest_rev = by_year.get(oldest["year"], {}).get("revenue")
        latest_rev = by_year.get(latest["year"], {}).get("revenue")
        if oldest_rev is not None and latest_rev is not None and oldest_rev > 0:
            rev_growth_pct = ((latest_rev - oldest_rev) / oldest_rev) * 100.0
            rev_growth_str = f"{rev_growth_pct:+.1f}%"
            rev_growth_cls = "positive" if rev_growth_pct >= 0 else "negative"
            rev_growth_sub = f"FY {oldest['year']} (${oldest_rev:,.0f}) → FY {latest['year']} (${latest_rev:,.0f})"
        else:
            rev_growth_str = "N/A"
            rev_growth_cls = ""
            rev_growth_sub = "Revenue data unavailable"

        # 3. Latest MD&A P/N Ratio
        latest_pn = latest.get("pn_ratio")
        latest_pn_str = f"{latest_pn:.2f}" if latest_pn is not None else "N/A"
        latest_pn_sub = "Positive-to-Negative Words"

        # 4. Latest Stock Price
        latest_sp_str = f"${latest_p:,.2f}" if latest_p is not None else "N/A"
        latest_sp_sub = f"FY {latest['year']} Balance Sheet Date"

        # Section: Executive Summary (Requirement 2)
        st.subheader("📌 Executive Summary")
        st.caption(f"Comprehensive executive snapshot for **{company_name}** across the selected {reporting_scope.lower()}.")

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
            <div class="metric-sub">{latest_sp_sub}</div>
        </div>
        """, unsafe_allow_html=True)

        c_p4.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">MD&A Sentiment</div>
            <div class="metric-name">Latest P/N Ratio</div>
            <div class="metric-val">{latest_pn_str}</div>
            <div class="metric-sub">{latest_pn_sub}</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"**Latest Fiscal Year Accounting Health (FY {latest['year']})**")
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)

        # 1. Current Ratio
        cr_val = f"{latest['current_ratio']:.2f}" if latest.get("current_ratio") is not None else "N/A"
        col_m1.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Liquidity</div>
            <div class="metric-name">Current Ratio</div>
            <div class="metric-val">{cr_val}</div>
            <div class="metric-sub">Current Assets / Liabilities</div>
        </div>
        """, unsafe_allow_html=True)

        # 2. Debt-to-Assets
        da_val = f"{latest['debt_to_assets'] * 100:.2f}%" if latest.get("debt_to_assets") is not None else "N/A"
        col_m2.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Solvency</div>
            <div class="metric-name">Debt-to-Assets</div>
            <div class="metric-val">{da_val}</div>
            <div class="metric-sub">Total Liabilities / Assets</div>
        </div>
        """, unsafe_allow_html=True)

        # 3. Total Asset Turnover
        at_val = f"{latest['asset_turnover']:.2f}" if latest.get("asset_turnover") is not None else "N/A"
        col_m3.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Efficiency</div>
            <div class="metric-name">Asset Turnover</div>
            <div class="metric-val">{at_val}</div>
            <div class="metric-sub">Revenue / Average Assets</div>
        </div>
        """, unsafe_allow_html=True)

        # 4. Net Profit Margin
        pm_val = f"{latest['profit_margin'] * 100:.2f}%" if latest.get("profit_margin") is not None else "N/A"
        col_m4.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Profitability</div>
            <div class="metric-name">Net Margin</div>
            <div class="metric-val">{pm_val}</div>
            <div class="metric-sub">Net Income / Revenue</div>
        </div>
        """, unsafe_allow_html=True)

        # Section: Historical Summary Table (Requirement 3)
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
        df_hist = pd.DataFrame(hist_summary_rows)
        st.dataframe(df_hist, width="stretch", hide_index=True)

        # Section: Interactive Dual-Axis Line Chart (Requirement 4)
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("📈 Stock Price vs. MD&A Sentiment P/N Ratio")
        st.caption("Chronological time-series comparison with dual Y-axes (Stock Price on Left, P/N Ratio on Right).")

        chrono_years = [str(r["year"]) for r in analysis_data]
        stock_series = [r.get("stock_price") for r in analysis_data]
        pn_series = [r.get("pn_ratio") for r in analysis_data]

        fig = make_subplots(specs=[[{"secondary_y": True}]])

        # Primary Axis: Stock Price ($)
        fig.add_trace(
            go.Scatter(
                x=chrono_years,
                y=stock_series,
                name="Stock Price ($)",
                mode="lines+markers",
                line=dict(color="#1F497D", width=3),
                marker=dict(size=8, symbol="circle"),
                hovertemplate="FY %{x}<br>Stock Price: $%{y:,.2f}<extra></extra>"
            ),
            secondary_y=False
        )

        # Secondary Axis: MD&A Sentiment P/N Ratio
        fig.add_trace(
            go.Scatter(
                x=chrono_years,
                y=pn_series,
                name="P/N Ratio",
                mode="lines+markers",
                line=dict(color="#E4572E", width=3, dash="dot"),
                marker=dict(size=8, symbol="diamond"),
                hovertemplate="FY %{x}<br>P/N Ratio: %{y:.2f}<extra></extra>"
            ),
            secondary_y=True
        )

        fig.update_layout(
            title_text=f"<b>{company_name} ({ticker}) — Stock Price vs. MD&A Sentiment P/N Ratio</b>",
            title_font=dict(size=16, color="#1F497D"),
            xaxis=dict(title="Fiscal Year", type="category", showgrid=True, gridcolor="#EDF2F7"),
            yaxis=dict(title="Historical Stock Price ($)", tickprefix="$", showgrid=True, gridcolor="#EDF2F7"),
            yaxis2=dict(title="MD&A Sentiment P/N Ratio", showgrid=False, tickformat=".2f"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            template="plotly_white",
            hovermode="x unified",
            height=460,
            margin=dict(l=60, r=60, t=80, b=50)
        )

        st.plotly_chart(fig, width="stretch")

        # Section: Optional Correlation Metric (Requirement 5)
        # Annual stock return: (Price_t / Price_{t-1}) - 1, paired with P/N Ratio_t
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

            *Note: The oldest fiscal year in the horizon has no prior baseline year and is excluded from YoY return pairing.*
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
                st.markdown("**Paired Annual Observations:**")
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

        # Section: Consolidated Data Sources & Methodology (Requirement 6)
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("📖 Data Sources & Methodology", expanded=False):
            st.markdown("""
            ### 🏛️ 1. SEC EDGAR Data Source & Filing Retrieval
            - **Primary Source**: U.S. Securities and Exchange Commission (SEC EDGAR) official submissions API (`https://data.sec.gov/submissions/`).
            - **Filing Form**: Form 10-K (Annual Reports filed by public entities under Section 13 or 15(d) of the Securities Exchange Act of 1934).
            - **Entity Lookup**: Official SEC Central Index Key (CIK) mapped directly from the company's exchange ticker symbol.
            - **Fair Access Compliance**: Declarative User-Agent contact header, polite pacing, and request rate throttling (maximum 10 requests per second) adhering strictly to SEC Fair Access regulations.

            ---
            ### 📅 2. Fiscal Year & `reportDate` Alignment
            - **Period End Alignment**: Each filing's financial statements are keyed to the official period-end date (`reportDate` / balance sheet date).
            - **Multi-Year Comparative Reporting**: Financial statements are aligned chronologically from oldest to newest fiscal year.
            - **Account Mapping**: Standard U.S. GAAP taxonomy concepts (Revenue, Net Income, Current Assets, Current Liabilities, Total Assets, Total Liabilities) are extracted directly from primary financial statements.

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
            - **Text Preprocessing**: Raw HTML/XBRL tags, tables, and scripts are stripped; punctuation and possessive suffixes (`'S`) are normalized; text is tokenized into uppercase alphabetic words without stemming.
            - **Dictionary Categories**: 8 domain-specific categories (*Positive, Negative, Uncertainty, Litigious, Strong Modal, Weak Modal, Constraining, Complexity*).
            - **Positive-to-Negative (P/N) Ratio**:
              $$\\text{P/N Ratio} = \\frac{\\text{Positive Words}}{\\text{Negative Words}}$$
              *(Protected against division-by-zero: if Negative = 0, returns Positive count if Positive > 0, else 1.0).*

            ---
            ### 🔢 5. Financial Ratio Formulas & Definitions
            *Standard financial-analysis ratios calculated using U.S. GAAP financial-statement data.*

            - **Current Ratio (Liquidity)**:
              $$\\text{Current Ratio} = \\frac{\\text{Current Assets}}{\\text{Current Liabilities}}$$
              *Evaluates short-term liquidity and ability to cover immediate liabilities.*
            - **Debt-to-Assets Ratio (Solvency)**:
              $$\\text{Debt-to-Assets} = \\frac{\\text{Total Liabilities}}{\\text{Total Assets}}$$
              *Measures capital structure leverage and the proportion of assets financed by creditors.*
            - **Total Asset Turnover (Efficiency)**:
              $$\\text{Total Asset Turnover} = \\frac{\\text{Revenue}}{\\text{Average Total Assets}}$$
              *Where Average Total Assets is calculated as:*
              $$\\text{Average Total Assets} = \\frac{\\text{Beginning Total Assets} + \\text{Ending Total Assets}}{2}$$
              *(Beginning assets are sourced from the comparative prior-year balance sheet).*
            - **Net Profit Margin (Profitability)**:
              $$\\text{Net Profit Margin} = \\frac{\\text{Net Income}}{\\text{Revenue}}$$
              *Measures operational efficiency and bottom-line earnings generated per dollar of revenue.*
            """)

    # --------------------------------------------------------------------------
    # TAB 2: FINANCIAL RATIOS
    # --------------------------------------------------------------------------
    with tab_ratios:
        st.subheader("Financial Ratios (Oldest to Newest)")

        ratio_table_data = []
        for r in analysis_data:
            ratio_table_data.append({
                "Fiscal Year": r["year"],
                "Current Ratio (Liquidity)": f"{r['current_ratio']:.2f}" if r["current_ratio"] is not None else "N/A",
                "Debt-to-Assets Ratio (Solvency)": f"{r['debt_to_assets'] * 100:.2f}%" if r["debt_to_assets"] is not None else "N/A",
                "Total Asset Turnover (Efficiency)": f"{r['asset_turnover']:.2f}" if r["asset_turnover"] is not None else "N/A",
                "Net Profit Margin (Profitability)": f"{r['profit_margin'] * 100:.2f}%" if r["profit_margin"] is not None else "N/A"
            })

        df_ratios = pd.DataFrame(ratio_table_data)
        st.dataframe(df_ratios, width="stretch", hide_index=True)

        st.markdown("""
        #### 📖 Ratio Formulas & Accounting Definitions
        - **Current Ratio (Liquidity)**: `Current Assets / Current Liabilities`  
          *Measures the firm's capacity to settle short-term obligations using liquid assets.*
        - **Debt-to-Assets Ratio (Solvency)**: `Total Liabilities / Total Assets`  
          *Measures financial leverage and the fraction of total assets funded by creditors rather than equity.*
        - **Total Asset Turnover (Efficiency)**: `Revenue / Average Total Assets`  
          *Measures efficiency in deploying total assets to generate top-line revenue.*
        - **Average Total Assets**: `(Beginning Total Assets + Ending Total Assets) / 2`  
          *Utilizes ending assets from the prior comparative period.*
        - **Net Profit Margin (Profitability)**: `Net Income / Revenue`  
          *Measures the percentage of each revenue dollar that flows to bottom-line net profit.*
        """)

        # Underlying Accounts Expander for Audit Verification
        with st.expander("🔍 View Underlying Financial Statement Accounts (Audit Verification)"):
            by_year = raw_accounts_info.get("by_year", {})
            accts_data = []
            for r in analysis_data:
                yr = r["year"]
                item = by_year.get(yr, {})
                accts_data.append({
                    "Fiscal Year": yr,
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
    # TAB 3: SENTIMENT ANALYSIS
    # --------------------------------------------------------------------------
    with tab_sentiment:
        st.subheader("Loughran-McDonald Financial Sentiment Analysis")
        st.caption("Word counts scored strictly against the official Loughran-McDonald Master Sentiment Dictionary.")

        other_cats = [c for c in available_cats if c not in ("Positive", "Negative")]

        sent_table_data = []
        for r in analysis_data:
            row_dict = {
                "Fiscal Year": r["year"],
                "Total Words": f"{r['total_words']:,d}" if r["total_words"] is not None else "N/A",
                "Positive": f"{r['positive']:,d}" if r["positive"] is not None else "N/A",
                "Negative": f"{r['negative']:,d}" if r["negative"] is not None else "N/A"
            }
            for cat in other_cats:
                val = r["other_sentiment"].get(cat)
                row_dict[cat] = f"{val:,d}" if val is not None else "N/A"
            row_dict["P/N Ratio"] = f"{r['pn_ratio']:.2f}" if r["pn_ratio"] is not None else "N/A"
            sent_table_data.append(row_dict)

        df_sentiment = pd.DataFrame(sent_table_data)
        st.dataframe(df_sentiment, width="stretch", hide_index=True)

        dict_word_count = results.get("dict_info", {}).get("total_words", 0)
        st.markdown(f"""
        **About the Sentiment Metrics**:
        - **Dictionary**: Loughran-McDonald Master Financial Sentiment Dictionary ({dict_word_count:,d} words).
        - **P/N Ratio**: `Positive Words / Negative Words`.
        - **Interpretation**: P/N > 1 means positive Loughran-McDonald word matches exceeded negative word matches. P/N < 1 means negative word matches exceeded positive word matches.
        """)

    # --------------------------------------------------------------------------
    # TAB 4: STOCK PRICES
    # --------------------------------------------------------------------------
    with tab_stock:
        st.subheader("Historical Stock Prices (Yahoo Finance)")
        st.caption("Split-adjusted closing prices matched to official SEC balance sheet dates.")

        # Sort stock_price_results chronologically for presentation
        chrono_stock = sorted(stock_price_results, key=lambda x: str(x["year"]))
        stock_table = []
        for s in chrono_stock:
            stock_table.append({
                "Fiscal Year": s["year"],
                "Balance Sheet Date": s["bs_date"],
                "Trading Date Used": s["trading_date"],
                "Stock Price": f"${s['price']:,.2f}" if s["price"] is not None else "N/A",
                "Date Adjustment": s.get("adjustment", "Exact Date"),
                "Status": s.get("status", "Retrieved")
            })

        df_stock = pd.DataFrame(stock_table)
        st.dataframe(df_stock, width="stretch", hide_index=True)

        st.info("""
        **Trading Day Methodology**:
        - **Source**: Yahoo Finance via `yfinance`.
        - **Prices**: Split-adjusted closing stock price (`Close`).
        - **Non-Trading Days**: When the fiscal year-end balance sheet date fell on a weekend or market holiday, the most recent previous trading day was automatically selected. Future trading dates are never used.
        """)

    # --------------------------------------------------------------------------
    # TAB 5: FINANCIAL STATEMENTS
    # --------------------------------------------------------------------------
    with tab_statements:
        st.subheader("Primary Financial Statements (Item 8)")
        st.caption("Extracted multi-year financial statements aligned across fiscal periods.")

        stmt_selector = st.radio(
            "Select Statement to View:",
            options=["Income Statement", "Balance Sheet", "Cash Flows"],
            horizontal=True
        )

        stmt_obj = statement_data.get(stmt_selector, {})
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
                # Dataframes do not render Markdown, so headers are shown in capitals
                row_dict = {"Line Item": label.upper() if is_header else f"    {label}"}
                row_vals = values.get(key, {})
                row_text = f"{key[0]} {label}".lower()
                for yr in years:
                    val = row_vals.get(yr)
                    if is_header or val is None:
                        row_dict[yr] = ""
                    elif "shares" in row_text:
                        # Share counts (e.g. 'Shares used in computing earnings
                        # per share') are not dollar amounts
                        row_dict[yr] = f"{val:,.0f}"
                    elif "per share" in row_text:
                        # Earnings / dividends per share keep their cents
                        row_dict[yr] = f"${val:,.2f}" if val >= 0 else f"(${-val:,.2f})"
                    else:
                        # Accounting style: negatives in parentheses
                        row_dict[yr] = f"${val:,.0f}" if val >= 0 else f"(${-val:,.0f})"
                stmt_rows.append(row_dict)

            df_stmt = pd.DataFrame(stmt_rows)
            st.dataframe(df_stmt, width="stretch", hide_index=True)
        else:
            st.warning(f"No statement data available for {stmt_selector}.")

    # --------------------------------------------------------------------------
    # TAB 6: DOWNLOAD REPORT
    # --------------------------------------------------------------------------
    with tab_download:
        st.subheader("📥 Download Master Financial Analysis Workbook")
        st.markdown(f"""
        Download the full Excel analysis workbook generated by the backend:
        - **File Name:** `{Path(output_filepath).name}`
        - **Worksheets Included:**
          1. `Income Statement` (Multi-year operations)
          2. `Balance Sheet` (Multi-year assets, liabilities, and equity)
          3. `Cash Flows` (Operating, investing, and financing flows)
          4. `Sentiment Analysis` (Loughran-McDonald category counts and P/N ratios)
          5. `Stock Prices` (Historical split-adjusted prices matched to balance sheet dates)
          6. `Analysis` (Chronological ratios, formula documentation, and dual-axis line chart)
        """)

        if Path(output_filepath).exists():
            with open(output_filepath, "rb") as f:
                excel_bytes = f.read()

            st.download_button(
                label="📥 Download Excel Analysis (.xlsx)",
                data=excel_bytes,
                file_name=Path(output_filepath).name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width="stretch"
            )
            st.success("Workbook is ready for download.")
        else:
            st.error(f"Workbook file '{output_filepath}' not found on disk.")

else:
    # Friendly status notice when selected company has not been analyzed yet (Requirement 4)
    if active_ticker == "AAPL":
        company_label = "Apple Inc. (AAPL)"
    elif active_ticker == "TSLA":
        company_label = "Tesla, Inc. (TSLA)"
    else:
        company_label = f"Company '{active_ticker}'"

    st.info(f"💡 **{company_label}** has not been analyzed in this session yet. Click **🚀 Run Analysis** in the sidebar to generate the report.")
