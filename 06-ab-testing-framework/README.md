# 🧪 Project 06 — A/B Testing Analysis Framework

> **A complete statistical A/B testing engine — runs hypothesis tests, calculates confidence intervals, determines significance, and exports a full diagnostic report automatically.**

---

## 🎯 What This Project Does

This framework analyses real A/B tests the way a data analyst would in a product or marketing team. Given any binary metric dataset (open/clicked/converted), it runs a full statistical analysis and produces a clear, interpretable verdict — backed by multiple statistical methods.

Three realistic tests are included out of the box:

| Test | Control | Variant | Metric |
|------|---------|---------|--------|
| Email Subject Line | "Your Weekly Summary" | "Don't Miss This Week's Highlights" | Open rate |
| Button Colour | Blue CTA | Green CTA | Click-through rate |
| Pricing Page | Original layout | Redesigned with social proof | Conversion rate |

---

## 📐 Statistical Methods Used

| Method | Purpose |
|--------|---------|
| **Two-proportion Z-test** | Primary hypothesis test |
| **Chi-squared test** | Independence test to confirm Z-test |
| **Wilson Score CI** | 95% Confidence interval per group |
| **Absolute Uplift** | Raw difference in conversion rate |
| **Relative Uplift** | % improvement of Variant over Control |
| **Statistical Power** | Probability the test detected a real effect |
| **Sample Size Calculator** | Pre-test planning tool |

---

## 🏆 Verdict Logic

```
p-value < 0.05  AND  uplift > 0  →  ✅  Variant Wins — Launch It
p-value < 0.05  AND  uplift < 0  →  ❌  Control Wins — Keep Original
p-value ≥ 0.05                   →  ⚠️   Inconclusive — Collect More Data
```

---

## 📋 Report Sheets

| Sheet | Contents |
|-------|----------|
| 📋 Executive Summary | All 3 tests side by side, colour-coded by outcome |
| Email Subject Line | Full stats, CI, daily trend for Test 1 |
| Landing Page Button | Full stats, CI, daily trend for Test 2 |
| Pricing Page Layout | Full stats, CI, daily trend for Test 3 |
| 🧮 Sample Size Calculator | Required users per group for any baseline + MDE combo |

---

## 📊 Sample Results

```
Email Subject Line Test
  Control: 21.36%  |  Variant: 26.60%  |  p=0.0000
  Verdict: ✅  Variant Wins — Launch It

Button Colour Test
  Control: 13.89%  |  Variant: 14.67%  |  p=0.5048
  Verdict: ⚠️   Inconclusive — Collect More Data

Pricing Page Test
  Control: 2.72%   |  Variant: 5.94%   |  p=0.0000
  Verdict: ✅  Variant Wins — Launch It
```

---

## 🚀 How to Run

### 1. Install dependencies
```bash
pip install pandas openpyxl scipy numpy
```

### 2. Generate test datasets
```bash
python generate_data.py
```

### 3. Run the full analysis
```bash
python ab_testing_framework.py
```

A timestamped `AB_Test_Report_*.xlsx` is created automatically.

---

## 📁 Project Structure

```
06-ab-testing-framework/
├── generate_data.py           ← Creates 3 A/B test datasets
├── test1_email.csv            ← Email open rate test (5,000 users)
├── test2_button.csv           ← Button click test (3,600 users)
├── test3_pricing.csv          ← Pricing conversion test (6,400 users)
├── ab_testing_framework.py    ← Main analysis engine
├── AB_Test_Report_*.xlsx      ← Output report (auto-generated)
└── README.md
```

---

## 💡 Key Design Decisions

| Decision | Reason |
|----------|--------|
| Dual Z-test + Chi-squared | Two independent methods build trust in the result |
| Wilson CI | More accurate than normal CI for small conversion rates |
| Inconclusive verdict | Prevents premature decisions on insufficient data |
| Sample size calculator | Pre-test planning is as critical as post-test analysis |
| Daily trend table | Detects novelty effects — early spikes that fade over time |

---

## 👤 Author

**Jatin Prasad**  
MS Robotics and Autonomous Systems (Systems Engineering)  
Ira A. Fulton Schools of Engineering, Arizona State University

*Part of a data analytics portfolio demonstrating statistical analysis and A/B testing skills.*
