import glob
import os
import time
import urllib.parse

import pandas as pd
from openpyxl.utils import get_column_letter

from nclt_utils import (
    CIN_PATTERN,
    LLPIN_PATTERN,
    configured_output_directory,
    extract_case_id_from_cells,
    extract_main_parties,
    extract_party_from_cells,
    extract_verified_id,
    find_latest_file,
    normalize_company_key,
    normalize_text,
)


AUDIT_SOURCE_PREFIX = "[AUDIT RECOVERED SOURCE:"
COURT_OPEN_MARKER = "\u25ba\u25ba\u25ba"
COURT_CLOSE_MARKER = "\u25c4\u25c4\u25c4"
RUN_TIMESTAMP = os.getenv("NCLT_RUN_TIMESTAMP", "").strip() or time.strftime(
    "%Y-%m-%d_%H-%M-%S"
)
OUTPUT_DIR = configured_output_directory()


def setup_stealth_driver():
    """Boot a headless Chrome browser for search-result lookups."""
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager

    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options,
    )


def get_latest_master_file():
    return find_latest_file("Consolidated_*Master_*.xlsx")


def clean_company_name(text):
    """Return the final main-case creditor and corporate debtor."""
    parsed = extract_main_parties(text)
    return parsed["creditor"] or None, parsed["debtor"] or None


def extract_id_from_text(text, company_name=None):
    """Extract a CIN/LLPIN, optionally requiring matching company context."""
    if company_name:
        return extract_verified_id(text, company_name)
    cin_match = CIN_PATTERN.search(text or "")
    if cin_match:
        return cin_match.group(1).upper()
    llp_match = LLPIN_PATTERN.search(text or "")
    if llp_match:
        return llp_match.group(1).upper()
    return None


def read_search_page(driver):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait

    wait = WebDriverWait(driver, 15)
    wait.until(
        lambda current: current.execute_script("return document.readyState") == "complete"
    )
    wait.until(lambda current: current.find_element(By.TAG_NAME, "body"))
    wait.until(
        lambda current: len(current.find_element(By.TAG_NAME, "body").text.strip()) > 20
    )
    return driver.find_element(By.TAG_NAME, "body").text


def fetch_cin_with_browser(driver, company_name):
    """Search for an ID and verify its result context against the company name."""
    if not company_name or len(company_name) < 3:
        return "N/A"

    try:
        queries = [
            f'"{company_name}" CIN ZaubaCorp',
            f'"{company_name}" CIN',
        ]
        for query_text in queries:
            query = urllib.parse.quote(query_text)
            driver.get(f"https://search.yahoo.com/search?p={query}")
            found_id = extract_verified_id(read_search_page(driver), company_name)
            if found_id:
                return found_id
        return "NOT FOUND"
    except Exception:
        return "ERROR"


def _court_name(row_text):
    court = row_text.replace(COURT_OPEN_MARKER, "")
    court = court.replace(COURT_CLOSE_MARKER, "")
    return court.strip()


def extract_master_records(dataframe):
    """Parse and deduplicate Master rows without opening a browser."""
    records = []
    current_court = "UNKNOWN COURT"
    seen_records = set()

    for index, row in dataframe.iterrows():
        cells = [str(cell) for cell in row if pd.notna(cell)]
        row_text = " ".join(cells)
        if not row_text:
            continue

        if row_text.startswith(AUDIT_SOURCE_PREFIX):
            source_file = row_text[len(AUDIT_SOURCE_PREFIX) :].rstrip("]").strip()
            current_court = f"AUDIT RECOVERY - {source_file}"
            continue

        if COURT_OPEN_MARKER in row_text:
            current_court = _court_name(row_text)
            continue

        parsed = extract_party_from_cells(cells)
        if not parsed:
            continue

        case_id = extract_case_id_from_cells(cells) or "UNKNOWN ID"
        debtor = parsed["debtor"]
        case_key = normalize_text(case_id)
        if case_id == "UNKNOWN ID":
            case_key = f"{normalize_text(current_court)} ROW {index}"
        record_key = (case_key, normalize_company_key(debtor))
        if record_key in seen_records:
            continue
        seen_records.add(record_key)

        records.append(
            {
                "Court": current_court,
                "Case ID": case_id,
                "Financial Creditor": parsed["creditor"],
                "Corporate Debtor": debtor,
                "_parse_confident": parsed["confident"],
                "_parse_reason": parsed["reason"],
            }
        )

    return records


def run_enrichment():
    print("==============================================")
    print("   NCLT CIN ENRICHER (BROWSER ENGINE v6.0)")
    print("==============================================\n")

    master_file = get_latest_master_file()
    if not master_file:
        raise RuntimeError("No Consolidated Master file found in this folder.")

    print(f"[*] Reading Master File: {master_file}")
    dataframe = pd.read_excel(master_file, header=None)
    parsed_records = extract_master_records(dataframe)
    print(f"[*] Parsed unique case/company rows: {len(parsed_records)}")

    searchable_count = sum(record["_parse_confident"] for record in parsed_records)
    driver = setup_stealth_driver() if searchable_count else None
    cin_cache = {}
    database = []

    try:
        for record in parsed_records:
            debtor = record["Corporate Debtor"]
            if not record["_parse_confident"]:
                cin = "PARSE REVIEW"
                extraction_status = record["_parse_reason"] or "Party parsing needs review"
                print(f"  > Parse review: {debtor[:55]} [{extraction_status}]")
            else:
                company_key = normalize_company_key(debtor)
                print(f"  > Scanning web for: {debtor[:45]}...", end=" ")
                if company_key not in cin_cache:
                    cin_cache[company_key] = fetch_cin_with_browser(driver, debtor)
                cin = cin_cache[company_key]
                extraction_status = "OK" if cin not in {"NOT FOUND", "ERROR"} else cin
                print(f"[{cin}]")

            database.append(
                {
                    "Court": record["Court"],
                    "Case ID": record["Case ID"],
                    "Financial Creditor": record["Financial Creditor"],
                    "Corporate Debtor": debtor,
                    "CIN / LLPIN": cin,
                    "Extraction Status": extraction_status,
                }
            )
    finally:
        if driver:
            driver.quit()

    if not database:
        print("\n[!] No company names found to extract.")
        return None

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_name = os.path.join(OUTPUT_DIR, f"CIN_Directory_{RUN_TIMESTAMP}.xlsx")
    export = pd.DataFrame(database)
    with pd.ExcelWriter(output_name, engine="openpyxl") as writer:
        export.to_excel(writer, index=False, sheet_name="CIN Directory")
        worksheet = writer.sheets["CIN Directory"]
        for index, column in enumerate(export.columns):
            max_length = max(export[column].astype(str).map(len).max(), len(column)) + 3
            worksheet.column_dimensions[get_column_letter(index + 1)].width = min(
                max_length,
                80,
            )

    print(f"\n[OK] Extracted {len(database)} unique case/company rows.")
    print(f"[OK] Saved database to: {output_name}")
    return output_name


if __name__ == "__main__":
    run_enrichment()
