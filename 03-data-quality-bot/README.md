# 🤖 Project 03 — Data Quality & Validation Bot

> **Point it at any CSV. It finds every data problem and tells you exactly what to fix.**

---

## 🎯 What This Project Does

This bot automatically scans any CSV dataset for data quality issues and produces a
detailed, styled Excel report — no manual inspection needed.

Drop in a CSV, run one command, and get back a full audit with:
- An overall **Health Score** for your dataset
- Every issue **categorised, counted, and flagged with severity**
- The exact **row numbers** of every problem record
- An **auto-cleaned version** of your data ready for analysis

---

## 🛠️ Tech Stack

| Tool | Purpose |
|------|---------|
| `pandas` | Data ingestion, profiling, and cleaning |
| `numpy` | Outlier detection via IQR method |
| `re` | Regex-based email validation |
| `openpyxl` | Styled Excel report generation |
| `Python 3.x` | Bot orchestration |

---

## 🔍 Checks Performed

| # | Check | Method | Severity |
|---|-------|--------|----------|
| 1 | **Missing Values** | `isnull()` scan across all columns | 🔴 High |
| 2 | **Duplicate Rows** | `duplicated()` full-row comparison | 🔴 High |
| 3 | **Negative Numerics** | Value < 0 in numeric columns | 🔴 High |
| 4 | **Invalid Emails** | Regex pattern validation | 🟡 Medium |
| 5 | **Age Violations** | Out-of-range check [0–120] | 🟡 Medium |
| 6 | **Casing Issues** | Whitespace + inconsistent capitalisation | 🟢 Low |
| 7 | **Invalid Dates** | `pd.to_datetime()` parse attempt | 🔴 High |
| 8 | **Numeric Outliers** | 3×IQR statistical fence method | 🟡 Medium |

---

## 📋 Output Sheets

| Sheet | Contents |
|-------|----------|
| **Health Dashboard** | KPI cards, health score, issue summary table, bar chart |
| **Missing Values** | Every null — row number and column name |
| **Duplicate Rows** | All duplicated records |
| **Negative Values** | Rows with negative numbers + exact values |
| **Invalid Emails** | Rows with malformed email addresses |
| **Age Violations** | Rows where age is outside [0–120] |
| **Casing Issues** | Text inconsistencies and whitespace problems |
| **Invalid Dates** | Rows with unparseable date strings |
| **Outliers** | Statistical outliers with fence values |
| **Cleaned Data** | Auto-fixed dataset ready for analysis |

---

## 🚀 How to Run

### 1. Install dependencies
```bash
pip install pandas openpyxl
```

### 2. Generate a sample dirty dataset (demo purposes)
```bash
python generate_dirty_data.py
```

### 3. Run the bot on any CSV
```bash
# On the provided sample dataset
python data_quality_bot.py dirty_data.csv

# On your own data
python data_quality_bot.py your_file.csv
```

A timestamped `DQ_Report_*.xlsx` is created automatically.

---

## 📁 Project Structure

```
03-data-quality-bot/
├── generate_dirty_data.py    ← Creates a 310-row dirty sample dataset
├── dirty_data.csv            ← Sample dataset with 8 types of planted issues
├── data_quality_bot.py       ← Main bot script
├── DQ_Report_*.xlsx          ← Output report (auto-generated)
└── README.md
```

---

## 🧹 Auto-Cleaning Applied

After flagging all issues, the bot produces a cleaned version of the data:

| Fix Applied | Method |
|-------------|--------|
| Removes duplicates | `drop_duplicates()` |
| Strips whitespace | `str.strip()` on all text columns |
| Standardises casing | `str.title()` on name/region/product/status |
| Removes negative rows | Drops rows where numeric columns < 0 |

---

## 💡 Sample Results (on dirty_data.csv)

```
310 rows × 10 columns scanned

Missing Values  → 23 issues   🔴
Duplicate Rows  → 14 issues   🔴
Negative Values →  9 issues   🔴
Invalid Emails  →  8 issues   🟡
Age Violations  →  5 issues   🟡
Invalid Dates   →  7 issues   🔴
Outliers        →  5 issues   🟡

Total: 71 issues found
Cleaned dataset: 294 rows retained
```

---

## 👤 Author

**Jatin Prasad**
MS Robotics and Autonomous Systems (Systems Engineering)
Ira A. Fulton Schools of Engineering, Arizona State University

*Part of a data analytics portfolio demonstrating automation, validation, and reporting skills.*
