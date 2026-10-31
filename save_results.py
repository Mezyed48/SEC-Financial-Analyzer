"""
Save finished analyses so the Streamlit app can show them as soon as it opens.

Runs the full Steps 2-9 pipeline for each company below and writes, into
'saved_results/':
  * <TICKER>_<N>.pkl                  - the analysis results the dashboard displays
  * <TICKER>_Financial_Analysis.xlsx  - the Excel workbook offered for download

Re-run this script whenever main.py changes so the saved results stay current:
    python3 save_results.py
"""

from datetime import date
import pickle
import shutil

from main import run_analysis_workflow, PROJECT_ROOT

SAVED_DIR = PROJECT_ROOT / "saved_results"
COMPANIES = ["AAPL", "TSLA"]
FILING_COUNT = 10


def save_company(ticker, filing_count):
    print(f"\nRunning {filing_count}-filing analysis for {ticker}...")
    results = run_analysis_workflow(ticker, filing_count, is_cli=False)
    if not results.get("success"):
        print(f"  [ERROR] {ticker}: {results.get('error')}")
        return False

    # Keep a copy of the workbook beside the results; the path is stored
    # relative to the project so it also works on Streamlit Cloud
    workbook_name = f"{ticker}_Financial_Analysis.xlsx"
    shutil.copy2(results["output_filepath"], SAVED_DIR / workbook_name)
    results["output_filepath"] = f"saved_results/{workbook_name}"

    # The full dictionary word lists are large and not needed for display
    dict_info = results.get("dict_info") or {}
    results["dict_info"] = {
        key: dict_info.get(key) for key in ("filename", "total_words", "categories")
    }
    results["saved_on"] = date.today().isoformat()

    out_path = SAVED_DIR / f"{ticker}_{filing_count}.pkl"
    with open(out_path, "wb") as f:
        pickle.dump(results, f)

    years = [row["year"] for row in results["analysis_data"]]
    print(f"  Saved {out_path.name} (FY {years[0]} - FY {years[-1]}) and {workbook_name}")
    return True


if __name__ == "__main__":
    SAVED_DIR.mkdir(exist_ok=True)
    ok = [save_company(ticker, FILING_COUNT) for ticker in COMPANIES]
    print(f"\nDone: {sum(ok)} of {len(ok)} companies saved to {SAVED_DIR}")
