<div align="center">
  
# 🧹 DataClean Pro
### Enterprise-Grade Data Processing & Analytics Dashboard

**Developed for the Data Analysis Internship at Thiranex**

[![Python 3.12+](https://img.shields.io/badge/python-3.12+-blue.svg)](https://www.python.org/downloads/release/python-3120/)
[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://streamlit.io)
[![Pandas](https://img.shields.io/badge/pandas-%23150458.svg?style=flat&logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![Testing](https://img.shields.io/badge/pytest-passing-success)](https://docs.pytest.org/en/7.4.x/)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

</div>

---

## 📖 Project Overview

DataClean Pro is an automated, end-to-end data preprocessing and reporting pipeline engineered during the **Thiranex Data Analysis Internship**. 

Real-world data is messy. This project solves the time-consuming problem of manual data cleaning by providing a highly robust, scalable Python engine that automatically detects and resolves data quality issues (missing values, duplicates, outliers, formatting inconsistencies). Furthermore, it democratizes this workflow by wrapping the engine in a **sleek, interactive web dashboard** allowing non-technical stakeholders to process data instantly.

---

## ✨ Key Features

- **🚀 Interactive Web Dashboard:** A stunning, dark-themed UI built with Streamlit for drag-and-drop file processing and live configuration.
- **📊 Dynamic Visual Summaries:** Embedded Plotly interactive charts that render data distributions on the fly before exporting.
- **🤖 Intelligent Automation:** Automatically imputes missing metrics using statistical medians, handles string standardization, parses datetimes, and flags statistical outliers using the IQR method.
- **📈 Automated Excel Reporting:** Not just a CSV dumper—generates a multi-sheet, heavily formatted `.xlsx` workbook complete with KPI dashboards, auto-filters, and native Excel charts.
- **🛡️ Enterprise-Grade Architecture:** Built with professional-grade Python standards, including strict type hinting, robust exception handling, and standard `logging` for audit trails.
- **🧪 Automated Test Suite:** Integrated `pytest` coverage verifying the mathematical and string manipulation integrity of the cleaning algorithms.

---

## 🛠️ Technology Stack

| Category | Tools / Libraries |
| :--- | :--- |
| **Core Engine** | Python, `pandas`, `numpy` |
| **Web Interface** | `streamlit` |
| **Data Visualization** | `plotly.express` |
| **Reporting** | `openpyxl` |
| **Testing** | `pytest` |

---

## 📁 Project Architecture

```text
DataClean-Pro/
├── .streamlit/
│   └── config.toml                # Custom dark theme configuration for the dashboard
├── data/
│   ├── raw/                       # Drop your dirty CSV/Excel files here
│   └── cleaned/                   # Output directory for processed CSV files
├── reports/                       # Output directory for formatted Excel dashboards
├── app.py                         # Main Streamlit Web Application
├── clean_and_report.py            # Core Data Engineering Engine (The Pipeline)
├── test_cleaning.py               # Pytest automated test suite
├── requirements.txt               # Project dependencies
└── README.md                      # Project documentation
```

---

## 🚀 Quick Start & Installation

### 1. Prerequisites
Ensure you have Python 3.10+ installed on your machine.

### 2. Setup Environment
Open your terminal (PowerShell, Command Prompt, or bash) and navigate to the project directory:
```bash
cd path/to/DataClean-Pro
```

Install the required dependencies:
```bash
pip install -r requirements.txt
```

---

## 💻 Usage Guide

You have two powerful ways to utilize DataClean Pro: the interactive Web Dashboard, or the automated Command Line Interface (CLI).

### Option A: The Web Dashboard (Recommended)
Launch the beautiful, user-friendly interface directly in your browser:
```bash
streamlit run app.py
```
1. Drag and drop any messy dataset (`.csv`, `.xlsx`).
2. Map your columns (Numeric, Text, Dates, Categories) using the smart sidebar dropdowns.
3. Click **"Clean Data & Generate Report"**.
4. Explore the **Visual Summaries** tab for live interactive charts.
5. Download your pristine CSV and formatted Excel Dashboard with one click.

### Option B: The Command Line (For Automation/Cron Jobs)
If you are integrating this into a larger automated pipeline, place your raw file into `data/raw/`, configure the constants at the top of `clean_and_report.py`, and run:
```bash
python clean_and_report.py
```
*Tip for Windows Users:* If you encounter character encoding issues with terminal emojis, run: `$env:PYTHONIOENCODING="utf-8"; python clean_and_report.py`

---

## 🧹 What Happens Under the Hood?

When the cleaning pipeline executes, it performs the following strict data governance checks:

| Problem Type | Algorithmic Solution |
| :--- | :--- |
| **Duplicates** | Identifies and drops 100% exact duplicate rows (keeps the first occurrence). |
| **Missing Metrics** | Imputes missing numerical data using the column's **Median** (resistant to outliers). |
| **Missing Text** | Flags and fills empty categorical/text cells with the string `"Unknown"`. |
| **Dirty Strings** | Strips leading/trailing whitespace, collapses internal double spaces, and enforces `Title Case`. |
| **Mixed Categories** | Standardizes categorical casing (e.g., merging "ELECTRONICS", "electronics", and " Electronics "). |
| **Date Parsing** | Coerces heterogeneous date strings into standard datetime objects, generating derived columns (`Month`, `Quarter`, `Day Of Week`). |
| **Outliers** | Statistically flags anomalies beyond $1.5 \times IQR$ into a boolean column, preserving data integrity while highlighting risks. |

---

## 🧪 Testing

To ensure the utmost reliability of the data processing algorithms, this project features an automated test suite.

Run the tests using:
```bash
pytest test_cleaning.py
```
*This validates missing value imputation, text standardization, duplicate removal, and column extraction functions.*

---

<div align="center">
  <b>Built with passion for Data Analytics excellence at Thiranex.</b>
</div>
