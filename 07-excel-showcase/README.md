# 📊 Project 07 — Excel Advanced Functions Showcase

> **A fully working Excel workbook demonstrating every advanced formula and feature used by professional Data Analysts — with live calculations, real validation rules, and conditional formatting.**

---

## 🎯 What This Project Does

This project generates a 9-sheet Excel workbook that serves as both a **skills demonstration** and a **reference guide**. Every yellow cell contains a real, working Excel formula that calculates live when opened. The workbook covers the full range of advanced Excel skills required for professional data analysis.

---

## 📋 Workbook Contents

| Sheet | Topic | Functions Demonstrated |
|-------|-------|----------------------|
| 🏠 Overview | Navigation guide | — |
| 📦 Product Lookup | Lookup & Reference | `XLOOKUP`, `INDEX-MATCH`, `IFERROR`, `VLOOKUP` |
| 📊 Sales Summary | Conditional Aggregation | `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `SUMPRODUCT` |
| 🔢 Nested Logic | Logical Functions | `Nested IF`, `IFS`, `SWITCH`, `CHOOSE`, `AND`, `OR` |
| 📅 Date Intelligence | Date & Time | `EOMONTH`, `NETWORKDAYS`, `DATEDIF`, `WEEKDAY`, `TODAY` |
| 📋 Dynamic Arrays | Modern Excel 365 | `UNIQUE`, `SORT`, `FILTER`, `SEQUENCE` |
| 🎨 Cond. Formatting | Visual Analytics | Colour scales, data bars, formula-based rules |
| ✅ Data Validation | Data Integrity | Dropdowns, numeric rules, range constraints |
| 📖 Formula Reference | Full cheatsheet | All functions with plain-English explanations |

---

## 🔑 Key Excel Features Demonstrated

### Lookup Functions
- **XLOOKUP** — modern replacement for VLOOKUP, searches in any direction
- **INDEX-MATCH** — two-function combo for flexible lookups (left-to-right and right-to-left)
- **IFERROR** — graceful error handling for lookup failures

### Conditional Aggregation
- **SUMIFS** — sum revenue matching multiple criteria (region AND status)
- **COUNTIFS** — count orders per rep and per status simultaneously
- **AVERAGEIFS** — average order value filtered by multiple conditions
- **SUMPRODUCT** — weighted calculations without helper columns

### Logic & Classification
- **Nested IF** — tiered grade assignment (A/B/C/D/F from score)
- **IFS** — cleaner alternative to nested IF for multiple conditions
- **SWITCH** — map short codes to full labels

### Date Intelligence
- **DATEDIF** — project duration in days/months/years
- **NETWORKDAYS** — working days excluding weekends
- **EOMONTH** — last day of any month
- **WEEKDAY + CHOOSE** — convert date to weekday name

### Dynamic Arrays (Excel 365)
- **UNIQUE** — extract distinct values automatically
- **FILTER** — return only rows matching a condition
- **SORT** — sort a dataset by any column
- **SEQUENCE** — generate number grids without dragging

### Conditional Formatting
- 3-colour scale on quarterly revenue (red → yellow → green)
- Data bars on annual totals
- Formula-based rule highlighting the top performer row in gold

### Data Validation
- Product and region dropdown lists
- Quantity constrained to whole numbers (1–100)
- Discount constrained to 0–0.5
- Custom error messages for every rule

---

## 🚀 How to Run

### 1. Install dependencies
```bash
pip install openpyxl
```

### 2. Generate the workbook
```bash
python build_showcase.py
```

`Excel_Advanced_Functions_Showcase.xlsx` is created in the same folder.

### 3. Open in Excel or LibreOffice
Click any **yellow cell** to see the formula in the formula bar.  
Try entering invalid values in the **✅ Data Validation** sheet to trigger error messages.

---

## 📁 Project Structure

```
07-excel-showcase/
├── build_showcase.py                  ← Generates the full workbook
├── Excel_Advanced_Functions_Showcase.xlsx  ← Output (auto-generated)
└── README.md
```

---

## 💡 Key Design Decisions

| Decision | Reason |
|----------|--------|
| Real Excel formulas (not just values) | Reviewer can open the file and see live calculations |
| Yellow highlighting for formula cells | Immediately identifies which cells contain formulas |
| 📖 Formula Reference sheet | Acts as a self-contained cheatsheet — the workbook explains itself |
| Colour-coded sheet tabs | Easier navigation across 9 sheets |
| Data Validation with error messages | Demonstrates production-level data integrity setup |

---

## 👤 Author

**Jatin Prasad**  
MS Robotics and Autonomous Systems (Systems Engineering)  
Ira A. Fulton Schools of Engineering, Arizona State University

*Part of a data analytics portfolio demonstrating advanced Excel skills for professional data analysis.*
