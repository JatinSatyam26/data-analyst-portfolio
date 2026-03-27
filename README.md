# 📊 Data Analyst Portfolio — Jatin Prasad

**MS Robotics & Autonomous Systems (Systems Engineering)**  
Ira A. Fulton Schools of Engineering — Arizona State University  

---

> This portfolio demonstrates end-to-end data analytics skills through 7 fully working projects — covering Excel automation, SQL analytics, data quality bots, scheduled reporting, interactive dashboards, A/B testing, and advanced Excel functions.
>
> Every project runs from a single command. No manual steps. No placeholder code.

---

## 🗂️ Projects

| # | Project | Core Skills | Tools |
|---|---------|-------------|-------|
| [01](#01) | [Automated Excel Report Generator](#01-automated-excel-report-generator) | Automation, Reporting | Python, pandas, openpyxl |
| [02](#02) | [SQL Customer Analytics](#02-sql-customer-analytics) | Advanced SQL, Reporting | SQLite, pandas, openpyxl |
| [03](#03) | [Data Quality & Validation Bot](#03-data-quality--validation-bot) | Bots, Automation, Validation | Python, pandas, numpy |
| [04](#04) | [Scheduled Email Report Bot](#04-scheduled-email-report-bot) | Bots, Scheduling, Automation | Python, smtplib, schedule |
| [05](#05) | [Interactive BI Dashboard](#05-interactive-bi-dashboard) | Dashboarding, BI, Visualization | Plotly Dash, pandas |
| [06](#06) | [A/B Testing Analysis Framework](#06-ab-testing-analysis-framework) | Statistical Analysis, Reporting | scipy, pandas, openpyxl |
| [07](#07) | [Excel Advanced Functions Showcase](#07-excel-advanced-functions-showcase) | Advanced Excel, Data Validation | openpyxl, Excel formulas |

---

## 01 — Automated Excel Report Generator

**What it does:** Ingests a raw CSV file, cleans the data, runs multi-dimensional analysis, and produces a fully formatted 6-sheet Excel report — with KPI cards, charts, and an audit trail. Zero manual steps.

**Skills:** Data cleaning · Aggregation · Excel automation · Report generation  
**Tools:** `pandas` · `openpyxl` · Python

**Sheets produced:** Executive Summary · Monthly Trends · Product Analysis · Regional Analysis · Sales Rep Leaderboard · Raw Data

```bash
cd 01-automated-excel-report
python generate_data.py      # create sample dataset
python report_generator.py   # generate full Excel report
```

📁 [View Project →](./01-automated-excel-report/)

---

## 02 — SQL Customer Analytics

**What it does:** Builds a relational SQLite database with 4 linked tables, runs 7 advanced SQL queries, and auto-exports all results to a formatted multi-sheet Excel report.

**SQL techniques:** CTEs · Window functions (`RANK`, `SUM OVER`, `AVG OVER`) · Correlated subqueries · `NOT EXISTS` · `CASE WHEN` · `PARTITION BY` · Multi-table JOINs

**Tools:** `sqlite3` · `pandas` · `openpyxl`

```bash
cd 02-sql-customer-analytics
python setup_database.py   # build the SQLite database
python run_analytics.py    # run all queries + export to Excel
```

📁 [View Project →](./02-sql-customer-analytics/)

---

## 03 — Data Quality & Validation Bot

**What it does:** Point it at any CSV file and it automatically runs 10 validation checks, scores the dataset out of 100, and exports a 9-sheet diagnostic Excel report with every issue flagged and colour-coded by severity.

**Checks:** Missing values · Duplicates · Outliers (IQR) · Invalid emails · Negative values · Future dates · Type mismatches · Casing inconsistencies

**Tools:** `pandas` · `numpy` · `openpyxl` · `re`

```bash
cd 03-data-quality-bot
python generate_messy_data.py   # create dirty sample dataset
python data_quality_bot.py      # run all checks + export report

# or scan your own file:
python data_quality_bot.py your_file.csv
```

📁 [View Project →](./03-data-quality-bot/)

---

## 04 — Scheduled Email Report Bot

**What it does:** Builds a daily sales Excel report and emails it as an attachment at a scheduled time every day — fully automated with logging, dry-run mode, and HTML email body.

**Features:** Daily scheduler · HTML email with KPI summary · Excel attachment · Run log · `--dry-run` mode for safe testing

**Tools:** `smtplib` · `schedule` · `pandas` · `openpyxl` · `logging`

```bash
cd 04-email-report-bot
python generate_data.py          # create sample data
python email_bot.py --dry-run    # test without sending email
python email_bot.py --now        # send immediately
python email_bot.py              # start daily scheduler
```

📁 [View Project →](./04-email-report-bot/)

---

## 05 — Interactive BI Dashboard

**What it does:** A fully interactive dark-themed BI dashboard in the browser. 5 live filters (date range, region, category, segment, status) update 5 KPI cards and 5 charts simultaneously in real time.

**Charts:** Monthly revenue trend (line + bar combo) · Order status donut · Revenue by product · Revenue by region · Sales rep leaderboard

**Tools:** `Plotly Dash` · `pandas`

```bash
cd 05-bi-dashboard
python generate_data.py   # create dataset
python dashboard.py       # launch dashboard
# Open http://127.0.0.1:8050 in your browser
```

📁 [View Project →](./05-bi-dashboard/)

---

## 06 — A/B Testing Analysis Framework

**What it does:** Loads A/B experiment data, runs a full suite of statistical tests, and produces a 5-sheet Excel report with significance verdicts and a clear SHIP IT / DO NOT SHIP recommendation.

**Tests:** Two-proportion Z-test · Welch's t-test · Cohen's h effect size · Statistical power · 95% confidence intervals · Device segmentation · Daily trend analysis

**Tools:** `scipy` · `pandas` · `numpy` · `openpyxl`

```bash
cd 06-ab-testing-framework
python generate_data.py           # simulate experiment data
python ab_testing_framework.py    # run full analysis + export report
```

**Sample result:** Treatment B — +23.1% conversion lift, +31.4% revenue lift, p < 0.01 → ✅ SHIP IT

📁 [View Project →](./06-ab-testing-framework/)

---

## 07 — Excel Advanced Functions Showcase

**What it does:** Generates a 9-sheet Excel workbook with live working formulas demonstrating every advanced Excel skill used by professional Data Analysts — plus a full formula reference cheatsheet.

**Functions covered:** `XLOOKUP` · `INDEX-MATCH` · `SUMIFS` · `COUNTIFS` · `AVERAGEIFS` · `SUMPRODUCT` · `Nested IF` · `IFS` · `SWITCH` · `DATEDIF` · `NETWORKDAYS` · `EOMONTH` · `UNIQUE` · `FILTER` · `SORT` · `SEQUENCE`

**Also includes:** Colour scale conditional formatting · Data bars · Formula-based highlight rules · Dropdown data validation · Numeric constraint validation

**Tools:** `openpyxl`

```bash
cd 07-excel-showcase
python build_showcase.py   # generate full Excel workbook
```

📁 [View Project →](./07-excel-showcase/)

---

## 🛠️ Tech Stack Summary

| Category | Tools |
|----------|-------|
| Languages | Python 3.x |
| Data Manipulation | pandas, numpy |
| SQL | SQLite, sqlite3 |
| Excel Automation | openpyxl |
| Dashboarding | Plotly Dash |
| Statistical Analysis | scipy |
| Scheduling & Bots | schedule, smtplib |
| Version Control | Git, GitHub |
| OS | Ubuntu (Linux) |

---

## 👤 About

**Jatin Prasad**  
MS Robotics and Autonomous Systems (Systems Engineering) — 2nd Semester  
Ira A. Fulton Schools of Engineering, Arizona State University  

**Leadership Background:**
- President, ECO Club FCRIT (2023–24) — led 28-member council, 500+ participant events
- Secretary, MESA FCRIT (2022–23) — founded IEI Student Chapter, launched Annual Technical Magazine *URJA 2022*, conducted 2 national-level competitions

📧 jsatyam@asu.edu  
🔗 [github.com/JatinSatyam26](https://github.com/JatinSatyam26)

---

*All projects were built on Ubuntu using Python 3. Each project folder contains a README with setup instructions and a detailed explanation of design decisions.*
