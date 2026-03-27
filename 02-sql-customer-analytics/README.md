# 🗄️ Project 02 — SQL Customer Analytics

> **Advanced SQL queries on a relational SQLite database — automatically exported to a formatted Excel report.**

---

## 🎯 What This Project Does

This project simulates a real business analytics workflow:

1. A **relational SQLite database** is built from scratch with 4 linked tables
2. A suite of **7 advanced SQL queries** is executed — each targeting a different business question
3. All results are **automatically exported** to a formatted, multi-sheet Excel report

No manual copy-pasting. Every query result lands in its own styled sheet with a title and description of the SQL technique used.

---

## 🛠️ Tech Stack

| Tool | Purpose |
|------|---------|
| `sqlite3` | Relational database engine |
| `pandas` | Query result handling and transformation |
| `openpyxl` | Excel report generation and styling |
| `Python 3.x` | Orchestration |

---

## 🗂️ Database Schema

```
customers       products
───────────     ─────────────
customer_id  ←─ product_id
name            product_name
email           category
region          unit_price
segment         cost_price
signup_date
age
     │
     ▼
   orders ──────────────── order_items
   ───────                 ────────────
   order_id ──────────────► order_id
   customer_id             product_id ──► products
   order_date              quantity
   status                  unit_price
   shipping_city           discount
```

- **200 customers** across 5 regions and 3 segments (Premium / Standard / Budget)
- **15 products** across 5 categories with cost and sale price
- **800 orders** spanning the full year 2023
- **2,000+ order line items**

---

## 📋 SQL Queries & Techniques

| Sheet | Business Question | SQL Techniques Used |
|-------|-------------------|---------------------|
| Revenue by Segment | Which customer segment generates the most revenue? | Multi-table JOIN, GROUP BY, aggregation |
| Top Customers by LTV | Who are the top 10 customers by lifetime value? | CTE, `RANK()` window function |
| Monthly Revenue Trend | How is revenue trending month over month? | CTE, `SUM OVER`, `AVG OVER` (rolling average) |
| Product Performance | Which products are most profitable? | JOIN, profit margin calculation |
| Inactive Customers | Which customers have never completed a purchase? | Correlated subquery, `NOT EXISTS`, `CASE WHEN` |
| Regional Ranking | How do segments rank within each region? | `RANK() OVER PARTITION BY` |
| Repeat vs One-Time Buyers | What share of customers are repeat buyers? | Subquery, `CASE WHEN`, `COUNT(*) OVER()` |

---

## 🚀 How to Run

### 1. Install dependencies
```bash
pip install pandas openpyxl
```

### 2. Build the database (first time only)
```bash
python setup_database.py
```

### 3. Run all queries and generate the report
```bash
python run_analytics.py
```

A timestamped `Customer_Analytics_Report_*.xlsx` is created automatically.

---

## 📁 Project Structure

```
02-sql-customer-analytics/
├── setup_database.py              ← Creates & populates customer_analytics.db
├── run_analytics.py               ← Runs all 7 SQL queries + exports to Excel
├── customer_analytics.db          ← SQLite database (auto-generated)
├── Customer_Analytics_Report_*.xlsx ← Output report (auto-generated)
└── README.md
```

---

## 💡 Key Design Decisions

| Decision | Reason |
|----------|--------|
| SQLite (no server needed) | Fully portable — runs anywhere without setup |
| Separated DB setup from analytics | Clean separation of concerns; DB built once, queried many times |
| SQL comments in every query | Explains the technique used — makes the code educational |
| Each query on its own Excel sheet | Easy navigation; each business question is self-contained |
| Profit margin calculated in SQL | Demonstrates ability to derive business metrics directly in queries |

---

## 👤 Author

**Jatin Prasad**  
MS Robotics and Autonomous Systems (Systems Engineering)  
Ira A. Fulton Schools of Engineering, Arizona State University

*Part of a data analytics portfolio demonstrating SQL, Python, and automated reporting skills.*
