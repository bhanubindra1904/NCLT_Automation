# NCLT Section 7 Automation

This project automates the identification of NCLT insolvency cases under
Section 7 of the Insolvency and Bankruptcy Code (IBC).

It collects cases from NCLT cause list PDFs, links companies to their CINs,
and prepares structured Excel reports.

# Workflow

1. `scraper.py` downloads the cause list PDFs for a selected date and extracts
   Section 7 cases.
2. `audit.py` checks the PDFs again using different extraction methods and
   identifies cases that may have been missed.
3. `cin_enricher.py` searches for the CIN or LLPIN of each company.
4. `debt_enricher.py` can be run separately to collect company status and the
   latest available financial and debt data from PrivateCircle.
5. Excel reports are saved in a dated folder inside `outputs/`.

# Main Files

`run_pipeline.py` runs the main NCLT, audit, and CIN steps in order.

`scraper.py` downloads and reads NCLT cause list PDFs.

`audit.py` rechecks PDFs and flags possible missed cases.

`cin_enricher.py` finds CIN or LLPIN details for companies.

`debt_enricher.py` collects PrivateCircle company and financial data.

`nclt_utils.py` contains shared PDF, text, case reference, and output helpers.

`requirements.txt` lists the Python packages required by the project.

`instructions.pdf` contains basic setup and usage instructions.

# Requirements

Windows 10 or 11

A supported 64 bit Python 3.x release

Google Chrome

Internet connection

PrivateCircle access for `debt_enricher.py`

Install the Python packages once from PowerShell:

```
pip install pandas openpyxl pdfplumber selenium webdriver_manager
```

# Run the Main Pipeline

Open PowerShell, move into the project folder, and run:

```
cd path\to\NCLT_Scraper
python run_pipeline.py
```

The scraper asks for the NCLT cause list date. Chrome may open during the
process because the project uses Selenium for browser automation.

# Run Debt Enrichment

Run this separately after the main pipeline has produced a CIN directory:

```
python debt_enricher.py
```

Log in to PrivateCircle when prompted. Do not commit usernames, passwords,
cookies, downloaded reports, or generated Excel files to GitHub.

# Output

Each run creates a timestamped folder inside `outputs/`. The folder contains
the downloaded PDFs, master case Excel files, audit results, and CIN results.
Debt enrichment Excel files are also saved there when `debt_enricher.py` is
run.

# Limitations

PDF layouts and OCR quality can affect extraction accuracy.

The audit identifies possible missed cases, but its review items may still
need manual checking.

CIN and financial data depend on the availability and accuracy of external
websites.

Selenium requires a working Chrome installation and internet connection.
