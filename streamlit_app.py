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

from pathlib import Path
import re
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# Import the verified backend pipeline from main.py without duplicating accounting logic
from main import run_analysis_workflow, PROJECT_ROOT


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
        border-radius: 4px;
        margin-bottom: 1.5rem;
        font-size: 0.95rem;
        color: #2D3748;
        line-height: 1.5;
    }
    .metric-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 1rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        text-align: center;
    }
    .metric-category {
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #718096;
        margin-bottom: 0.25rem;
    }
    .metric-name {
        font-size: 0.9rem;
        font-weight: 600;
        color: #2D3748;
        margin-bottom: 0.4rem;
    }
    .metric-val {
        font-size: 1.6rem;
        font-weight: 700;
        color: #1F497D;
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
# SESSION STATE INITIALIZATION
# ==============================================================================

if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None

if "last_ticker" not in st.session_state:
    st.session_state.last_ticker = ""


# ==============================================================================
# SIDEBAR / USER INPUT CONTROLS
# ==============================================================================

with st.sidebar:
    st.header("⚙️ Analysis Settings")
    st.markdown("Configure target public company and historical period horizon:")

    input_ticker = st.text_input(
        "Ticker Symbol:",
        value=st.session_state.last_ticker or "AAPL",
        max_chars=10,
        help="Enter a public company ticker symbol listed on a U.S. exchange (e.g., AAPL, MSFT, NVDA)."
    ).strip().upper()

    input_filing_count = st.number_input(
        "Number of Recent 10-K Filings:",
        min_value=1,
        max_value=10,
        value=4,
        step=1,
        help="Select between 1 and 10 fiscal years (default is 4; final run is 10)."
    )

    run_btn = st.button("🚀 Run Analysis", type="primary", width="stretch")

    st.markdown("---")
    st.caption("🔒 **SEC Fair Access Compliant**: Requests are throttled and use declarative User-Agent contact credentials.")
    st.caption("📚 **Accounting Standards**: Follows U.S. GAAP ratios and Loughran-McDonald Financial Sentiment Dictionary.")


# ==============================================================================
# RUN ANALYSIS TRIGGER (RERUN-SAFE)
# ==============================================================================

if run_btn:
    # Input validation
    clean_ticker = input_ticker.strip().upper()
    if not clean_ticker or not re.match(r"^[A-Z0-9\.\-]+$", clean_ticker):
        st.error("❌ Please enter a valid stock ticker symbol (letters, numbers, dots, or hyphens only).")
    else:
        st.session_state.last_ticker = clean_ticker

        # Progress tracking status container
        status_box = st.status(f"Starting SEC 10-K Analysis for **{clean_ticker}** ({input_filing_count} filings)...", expanded=True)
        progress_bar = st.progress(0.0)

        def on_progress(step_num, stage_name, description):
            pct = min(1.0, step_num / 10.0)
            progress_bar.progress(pct)
            status_box.write(f"**Step {step_num}/10: {stage_name}** — {description}")

        try:
            # Execute backend workflow without duplicating logic
            results = run_analysis_workflow(
                ticker=clean_ticker,
                filing_count=input_filing_count,
                progress_callback=on_progress,
                is_cli=False
            )

            if results.get("success"):
                progress_bar.progress(1.0)
                status_box.update(label=f"✅ Analysis for {clean_ticker} complete!", state="complete", expanded=False)
                st.session_state.analysis_results = results
            else:
                progress_bar.empty()
                status_box.update(label=f"❌ Analysis failed for {clean_ticker}", state="error", expanded=True)
                st.error(f"**Error:** {results.get('error')}")
                st.session_state.analysis_results = None

        except Exception as ex:
            progress_bar.empty()
            status_box.update(label="❌ Unexpected Error", state="error", expanded=True)
            st.error(f"An unexpected error occurred during execution: {ex}")
            st.session_state.analysis_results = None


# ==============================================================================
# DISPLAY RESULTS DASHBOARD
# ==============================================================================

results = st.session_state.analysis_results

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

    # Company Header Banner
    st.markdown(f"## 🏢 {company_name} ({ticker})")
    badge_col1, badge_col2, badge_col3 = st.columns(3)
    badge_col1.info(f"**SEC CIK:** `{cik}`")
    badge_col2.info(f"**Fiscal Horizon:** `{num_years} Fiscal Years` ({analysis_data[0]['year']} – {analysis_data[-1]['year']})")
    badge_col3.info(f"**Excel Master File:** `{Path(output_filepath).name}`")

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
        latest = analysis_data[-1]  # Most recent fiscal year
        st.subheader(f"Latest Fiscal Year Highlights — FY {latest['year']}")

        col_m1, col_m2, col_m3, col_m4, col_m5, col_m6 = st.columns(6)

        # 1. Current Ratio
        cr_val = f"{latest['current_ratio']:.2f}" if latest.get("current_ratio") is not None else "N/A"
        col_m1.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Liquidity</div>
            <div class="metric-name">Current Ratio</div>
            <div class="metric-val">{cr_val}</div>
        </div>
        """, unsafe_allow_html=True)

        # 2. Debt-to-Assets
        da_val = f"{latest['debt_to_assets'] * 100:.2f}%" if latest.get("debt_to_assets") is not None else "N/A"
        col_m2.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Solvency</div>
            <div class="metric-name">Debt-to-Assets</div>
            <div class="metric-val">{da_val}</div>
        </div>
        """, unsafe_allow_html=True)

        # 3. Total Asset Turnover
        at_val = f"{latest['asset_turnover']:.2f}" if latest.get("asset_turnover") is not None else "N/A"
        col_m3.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Efficiency</div>
            <div class="metric-name">Asset Turnover</div>
            <div class="metric-val">{at_val}</div>
        </div>
        """, unsafe_allow_html=True)

        # 4. Net Profit Margin
        pm_val = f"{latest['profit_margin'] * 100:.2f}%" if latest.get("profit_margin") is not None else "N/A"
        col_m4.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Profitability</div>
            <div class="metric-name">Net Margin</div>
            <div class="metric-val">{pm_val}</div>
        </div>
        """, unsafe_allow_html=True)

        # 5. P/N Ratio
        pn_val = f"{latest['pn_ratio']:.2f}" if latest.get("pn_ratio") is not None else "N/A"
        col_m5.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">MD&A Sentiment</div>
            <div class="metric-name">P/N Ratio</div>
            <div class="metric-val">{pn_val}</div>
        </div>
        """, unsafe_allow_html=True)

        # 6. Stock Price
        sp_val = f"${latest['stock_price']:,.2f}" if latest.get("stock_price") is not None else "N/A"
        col_m6.markdown(f"""
        <div class="metric-card">
            <div class="metric-category">Market Value</div>
            <div class="metric-name">Stock Price</div>
            <div class="metric-val">{sp_val}</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Interactive Dual-Axis Line Chart: Stock Price ($) vs. MD&A Sentiment P/N Ratio
        st.subheader("Stock Price vs. MD&A Sentiment P/N Ratio")
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
        - **Interpretation**: Values > 1.0 indicate predominantly optimistic tone; values < 1.0 reflect cautious tone typical of SEC risk disclosures.
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
    # Initial state before running analysis
    st.info("👈 Enter a ticker symbol in the sidebar and click **Run Analysis** to retrieve SEC filings and generate reports.")
