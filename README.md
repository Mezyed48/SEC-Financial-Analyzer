# SEC 10-K Financial & Sentiment Analyzer

## Overview

The **SEC 10-K Financial & Sentiment Analyzer** is an end-to-end Python accounting and financial analytics system designed to automate multi-year corporate performance evaluation. The application:

- Retrieves official Form 10-K annual filings directly from SEC EDGAR.
- Extracts plain-text Management's Discussion & Analysis (Item 7 MD&A) sections.
- Extracts primary financial statements (Item 8 Income Statement, Balance Sheet, and Statement of Cash Flows).
- Cleans and normalizes MD&A text for financial vocabulary preservation.
- Performs financial sentiment analysis using the domain-specific Loughran-McDonald Master Sentiment Dictionary.
- Retrieves historical split-adjusted closing stock prices from Yahoo Finance matched to official balance sheet dates.
- Calculates key financial ratios assessing liquidity, solvency, efficiency, and profitability.
- Generates a professionally formatted multi-worksheet Excel financial model (`output/<TICKER>_Financial_Analysis.xlsx`) complete with an embedded dual-axis combination line chart.
- Provides both a Command Line Interface (CLI) and an interactive local Streamlit web application.

---

## Features (Steps 2 through 9 Workflow)

- **Step 2 — User Input Validation**: Validates ticker symbols and filing horizons (1 to 10 fiscal years).
- **Step 3 — SEC EDGAR Retrieval**: Automates CIK lookup, rate limiting (max 10 requests/second), and primary Form 10-K HTML downloads.
- **Step 4 — MD&A Extraction**: Isolates Item 7 MD&A plain text while removing tables, scripts, and non-narrative artifacts.
- **Step 5 — Financial Statements Extraction**: Parses Item 8 tables for operations, balance sheets, and cash flows; merges multiple filings into a single comparative multi-year workbook.
- **Step 6 — Text Cleaning & Normalization**: Strips repeated SEC page headers/footers, converts curved punctuation, and preserves essential negative qualifiers (*not*, *no*, *never*, *without*).
- **Step 7 — Loughran-McDonald Sentiment Scoring**: Counts frequencies across 8 financial sentiment categories (*Positive, Negative, Uncertainty, Litigious, Strong Modal, Weak Modal, Constraining, Complexity*) and computes zero-division protected Positive-to-Negative (P/N) ratios.
- **Step 8 — Stock Price Retrieval**: Retrieves split-adjusted closing prices from Yahoo Finance as of the company's fiscal year-end balance sheet date (`reportDate`), automatically rolling back to the preceding trading day for non-trading dates.
- **Step 9 — Financial Ratios & Dual-Axis Visualization**: Computes four accounting ratios, arranges tables in chronological order (oldest to newest), formats cells with plain-text formula explanations, and embeds an openpyxl dual-axis Line Chart (Stock Price vs. P/N Ratio).

---

## Financial Ratios

The program calculates four core financial ratios for every fiscal year:

1. **Liquidity — Current Ratio**:
   $$\text{Current Ratio} = \frac{\text{Current Assets}}{\text{Current Liabilities}}$$
   *Measures the firm's capacity to cover short-term debt obligations with liquid assets.*

2. **Solvency — Debt-to-Assets Ratio**:
   $$\text{Debt-to-Assets Ratio} = \frac{\text{Total Liabilities}}{\text{Total Assets}}$$
   *Measures long-term financial leverage and the fraction of total capital provided by creditors.*

3. **Efficiency — Total Asset Turnover**:
   $$\text{Total Asset Turnover} = \frac{\text{Revenue or Net Sales}}{\text{Average Total Assets}}$$
   $$\text{Average Total Assets} = \frac{\text{Beginning Total Assets} + \text{Ending Total Assets}}{2}$$
   *Evaluates how efficiently the company deploys assets to generate top-line revenue. Prior-year beginning assets for the oldest period are retrieved from comparative balance sheet tables in the oldest filing.*

4. **Profitability — Net Profit Margin**:
   $$\text{Net Profit Margin} = \frac{\text{Net Income}}{\text{Revenue or Net Sales}}$$
   *Measures the proportion of each dollar of revenue converted into bottom-line profit.*

---

## Project Structure

```text
├── main.py                     # Core backend executing Steps 2 through 9
├── streamlit_app.py            # Local Streamlit web dashboard
├── requirements.txt            # Python dependencies
├── .env.example                # Template for SEC User-Agent credentials
├── .gitignore                  # Git exclusion rules (protects private data)
├── README.md                   # Project documentation
│
├── dictionaries/               # Sentiment dictionary directory (contains .gitkeep)
│   └── Loughran-McDonald_MasterDictionary_1993-2025.csv  (user-supplied)
│
├── filings/                    # Downloaded Form 10-K HTML filings (local, Git-ignored)
│   └── <TICKER>/
│
├── extracted_text/             # Raw extracted MD&A text files (local, Git-ignored)
│   └── <TICKER>/
│
├── cleaned_text/               # Normalized MD&A text files (local, Git-ignored)
│   └── <TICKER>/
│
└── output/                     # Generated master Excel workbooks (local, Git-ignored)
    └── <TICKER>_Financial_Analysis.xlsx
```

---

## Installation

Clone the repository and install required Python packages:

```bash
git clone https://github.com/Mezyed48/SEC-Financial-Analyzer.git
cd SEC-Financial-Analyzer
python3 -m pip install -r requirements.txt
```

---

## Configuration & SEC User-Agent

SEC EDGAR requires all programmatic requests to declare a User-Agent containing user identification and contact information.

1. Copy the example configuration file:
   ```bash
   cp .env.example .env
   ```
2. Open `.env` and fill in your details:
   ```env
   SEC_USER_NAME=Your Name
   SEC_USER_ORGANIZATION=Your University or Organization
   SEC_USER_EMAIL=your_email@example.com
   ```
*(Note: `.env` is excluded by `.gitignore` so your personal email and contact information will never be uploaded to GitHub.)*

---

## Loughran-McDonald Dictionary Setup

Because of file size and licensing guidelines, third-party sentiment dictionaries are not bundled in this repository.

1. Download the official **Loughran-McDonald Master Dictionary** (CSV or Excel) from the Notre Dame Software Repository for Accounting and Finance (SRAF):  
   [https://sraf.nd.edu/loughranmcdonald-master-dictionary/](https://sraf.nd.edu/loughranmcdonald-master-dictionary/)
2. Place the downloaded dictionary file into the `dictionaries/` folder:
   ```text
   dictionaries/Loughran-McDonald_MasterDictionary_1993-2025.csv
   ```
The application dynamically detects and loads any valid Loughran-McDonald CSV or Excel file in that folder.

---

## Running the Application

### Option A: Command Line Interface (CLI)

Run the program from your terminal:
```bash
python3 main.py
```
- Enter the public company ticker symbol (e.g., `AAPL`, `MSFT`, `NVDA`).
- Enter the number of recent 10-K filings to retrieve (between `1` and `10`).

The terminal outputs real-time status reports, financial statement line items, sentiment scores, and full underlying account verification tables.

### Option B: Local Streamlit Web Application

To launch the interactive graphical web dashboard:
```bash
streamlit run streamlit_app.py
```
This opens a local browser interface featuring:
- 6 interactive KPI cards for latest-year metrics.
- Plotly interactive dual-axis chart (Historical Stock Price vs. MD&A Sentiment P/N Ratio).
- Multi-year comparative financial statement viewers.
- Direct Excel workbook download button.

---

## Output

Completed Excel workbooks are saved automatically in the `output/` directory:
```text
output/<TICKER>_Financial_Analysis.xlsx
```
The workbook contains six distinct worksheets:
1. `Income Statement`
2. `Balance Sheet`
3. `Cash Flows`
4. `Sentiment Analysis`
5. `Stock Prices`
6. `Analysis` (Ratios, plain-text formula documentation, and dual-axis combination chart)

---

## Limitations

- Designed specifically for U.S. domestic operating companies filing annual reports on **Form 10-K**. Foreign private issuers filing Form 20-F or 40-F use different reporting frameworks and are not supported.
- Financial statement disclosure structures and line item labels can vary substantially across different industries and issuers (e.g., banking vs. software). The program prioritizes standard GAAP presentation patterns.
