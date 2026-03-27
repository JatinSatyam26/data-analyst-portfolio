# 📊 Project 05 — Interactive BI Dashboard

> **A fully interactive browser-based Business Intelligence dashboard built with Plotly Dash. All charts update in real time when filters change.**

---

## 🎯 What This Project Does

This project replicates the kind of live BI dashboard you'd find in tools like Tableau or Power BI — built entirely in Python. It loads a 1,200-row sales dataset and renders an interactive multi-chart dashboard in your browser. Every filter change instantly updates all 5 charts and all 5 KPI cards simultaneously.

---

## 🛠️ Tech Stack

| Tool | Purpose |
|------|---------|
| `Dash` (Plotly) | Web application framework for the dashboard |
| `Plotly` | Interactive chart rendering |
| `pandas` | Data loading, filtering, aggregation |
| `Python 3.x` | Backend logic and callbacks |

---

## 📊 Dashboard Components

### 5 KPI Cards (top row)
| Card | Metric |
|------|--------|
| 💰 Total Revenue | Sum of completed order revenue |
| 📦 Total Orders | All orders in filtered range |
| 🧾 Avg Order Value | Mean revenue per completed order |
| ↩️ Return Rate | % of orders with Returned status |
| ✅ Completed Orders | Count of completed orders |

### 5 Interactive Charts
| Chart | Type | Shows |
|-------|------|-------|
| Monthly Revenue Trend | Line + Bar combo | Revenue line + order volume bars per month |
| Order Status | Donut | Breakdown of Completed / Pending / Returned / Cancelled |
| Revenue by Product | Horizontal bar | Products ranked by revenue |
| Revenue by Region | Pie | Regional revenue share |
| Sales Rep Leaderboard | Horizontal bar | Reps ranked by revenue, top rep highlighted |

### 5 Filters (all charts respond instantly)
- **Date Range** — pick any date window
- **Region** — multi-select (North / South / East / West)
- **Category** — Electronics vs Accessories
- **Segment** — Premium / Standard / Budget
- **Order Status** — Completed / Pending / Returned / Cancelled

---

## 🚀 How to Run

### 1. Install dependencies
```bash
pip install pandas plotly dash
```

### 2. Generate sample data (first time only)
```bash
python generate_data.py
```

### 3. Launch the dashboard
```bash
python dashboard.py
```

### 4. Open your browser
```
http://127.0.0.1:8050
```

Press `Ctrl+C` in the terminal to stop.

---

## 📁 Project Structure

```
05-bi-dashboard/
├── generate_data.py      ← Creates 1,200-row dashboard_data.csv
├── dashboard_data.csv    ← Auto-generated sales dataset
├── dashboard.py          ← Main dashboard app
└── README.md
```

---

## 💡 Key Design Decisions

| Decision | Reason |
|----------|--------|
| Dark theme UI | Matches professional BI tool aesthetics (Tableau Dark, Power BI) |
| Dual-axis revenue trend | Shows both revenue value and order volume in one chart |
| Single Dash callback | All 5 charts update atomically from one filter change — no partial updates |
| Top rep highlighted in gold | Immediate visual identification of the leaderboard winner |
| `debug=False` in production | Prevents hot-reload in deployed environments |

---

## 👤 Author

**Jatin Prasad**  
MS Robotics and Autonomous Systems (Systems Engineering)  
Ira A. Fulton Schools of Engineering, Arizona State University

*Part of a data analytics portfolio demonstrating interactive dashboarding and BI skills.*
