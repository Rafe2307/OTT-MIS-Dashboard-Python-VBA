# 📊 OTT Platform — Sales & Revenue MIS Dashboard Generator
### Python + VBA Automation | Abdul Jabbar M

![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python)
![VBA](https://img.shields.io/badge/VBA-Excel-green?logo=microsoftexcel)
![Excel](https://img.shields.io/badge/Output-Excel_Dashboard-217346?logo=microsoftexcel)
![Status](https://img.shields.io/badge/Status-Production_Ready-brightgreen)

---

## 🎯 What This Project Does

This project **automatically generates a fully formatted, multi-sheet Excel MIS Dashboard** from raw OTT platform data (Amazon Prime & Netflix) — replacing hours of manual work with a single script execution.

**Two implementations included:**
- 🐍 **Python version** — for scheduled/automated runs, large datasets, server environments
- 📋 **VBA version** — for Excel-native automation, one-click dashboard generation

---

## 📁 Project Structure

```
OTT-MIS-Dashboard/
│
├── 📄 OTT_MIS_Dashboard_Generator.py   # Python automation script
├── 📄 OTT_MIS_Dashboard_VBA.bas        # VBA macro (import into Excel)
├── 📄 amazon_prime_titles.csv          # Sample input data
├── 📄 netflix_titles.csv               # Sample input data
└── 📄 README.md                        # This file
```

---

## 📊 Dashboard Output — 6 Sheets (Python) / 5 Sheets (VBA)

| Sheet | Contents |
|---|---|
| **1. Executive Summary** | KPI cards — Total Titles, Platform share, Movies, TV Shows, Countries |
| **2. Platform Comparison** | Amazon Prime vs Netflix head-to-head metrics |
| **3. Genre Analysis** | Top 15 genres by platform and combined volume |
| **4. Country Report** | Top 10 content-producing countries |
| **5. YoY Trend** | Year-on-year content growth with % change |
| **6. Raw Data** | Cleaned, consolidated 18,471-record dataset |

---

## ⚡ Key Features

- ✅ **Reads raw CSV data** — no manual data prep needed
- ✅ **Auto-calculates KPIs** — totals, shares, YoY growth, rankings
- ✅ **Fully formatted output** — colour-coded headers, alternating rows, bold totals
- ✅ **Platform colour-coding** — Amazon (cyan), Netflix (red) throughout
- ✅ **Insight generation** — auto-written business insights in each sheet
- ✅ **Error handling** — validates files exist, handles missing data gracefully
- ✅ **Progress logging** — shows step-by-step status in console (Python)
- ✅ **One-click run** — single command / single macro call

---

## 🚀 How to Run

### Python Version
```bash
# 1. Install dependencies
pip install pandas openpyxl xlsxwriter

# 2. Place CSV files in same folder as script

# 3. Run
python OTT_MIS_Dashboard_Generator.py

# Output: OTT_MIS_Dashboard_YYYY-MM-DD.xlsx
```

### VBA Version
```
1. Open Microsoft Excel
2. Press Alt + F11 to open VBA Editor
3. Insert > Module
4. Paste contents of OTT_MIS_Dashboard_VBA.bas
5. Create a sheet named "Raw_Data" with your data
6. Press Alt + F8 > Select "GenerateMISDashboard" > Run
```

---

## 📈 Sample Results (from 18,471 titles)

| Metric | Value |
|---|---|
| Total Titles Processed | 18,471 |
| Amazon Prime Titles | 9,668 (52.3%) |
| Netflix Titles | 8,806 (47.7%) |
| Countries Covered | 88 |
| Unique Genres | 64 |
| Processing Time (Python) | ~8 seconds |
| Output File Size | ~1 MB |

---

## 🛠️ Technical Stack

**Python:**
- `pandas` — data loading, cleaning, aggregation
- `xlsxwriter` — Excel file creation with formatting
- `openpyxl` — Excel read/write support
- `datetime` — dynamic date stamping

**VBA:**
- Native Excel VBA (no external libraries needed)
- Works on Excel 2016, 2019, 2021, Office 365

---

## 💡 Business Impact

This type of automation was built to solve real reporting challenges:

- **Before:** 2–3 hours of manual data extraction, pivot tables, formatting
- **After:** Under 10 seconds — fully formatted, ready-to-present dashboard
- **Accuracy:** 100% — eliminates manual copy-paste errors
- **Reusability:** Drop new CSV files → run script → fresh dashboard instantly

---

## 👤 Author

**Abdul Jabbar M**
- 🔗 [LinkedIn](https://www.linkedin.com/in/abdul-jabbar-m/)
- 💼 Deputy Manager — Wipro Technologies
- 🛠️ Data Analyst | Power BI Developer | VBA & Python Automation Specialist
- 📧 abduljabbarrafe.m@gmail.com

---

## 📌 Related Skills Demonstrated

`Python` `VBA Macros` `Excel Automation` `MIS Reporting` `ETL` `Data Analytics`
`Dashboard Development` `Report Automation` `Data Visualization` `Business Intelligence`

---

*⭐ If this project helped you, consider giving it a star on GitHub!*
