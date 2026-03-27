# 📧 Project 04 — Scheduled Email Report Bot

> **Automatically builds a daily sales report and emails it as a formatted Excel attachment — every day, zero human involvement.**

---

## 🎯 What This Project Does

This bot runs on a daily schedule. At a configured time each morning it:

1. Pulls the latest data from `sales_data.csv`
2. Builds a fresh, formatted 4-sheet Excel report
3. Composes a professional HTML email with a KPI summary table
4. Attaches the Excel report and sends it to all configured recipients
5. Logs every run with timestamp and outcome to `bot_log.txt`

Set it up once. It runs every day automatically.

---

## 🛠️ Tech Stack

| Tool | Purpose |
|------|---------|
| `smtplib` | SMTP email sending (built-in Python) |
| `email` (MIME) | HTML email + Excel attachment composition |
| `schedule` | Daily job scheduling |
| `pandas` | Data loading and aggregation |
| `openpyxl` | Excel report generation |
| `logging` | Run history and error tracking |

---

## 📁 Project Structure

```
04-email-report-bot/
├── generate_data.py      ← Creates sample sales_data.csv
├── sales_data.csv        ← Auto-generated sales dataset
├── report_builder.py     ← Builds the formatted Excel report
├── email_bot.py          ← Main bot: scheduler + email sender
├── bot_log.txt           ← Auto-generated run log
└── README.md
```

---

## ⚙️ Configuration

Before running, edit the `CONFIG` block at the top of `email_bot.py`:

```python
CONFIG = {
    "SENDER_EMAIL":    "your_email@gmail.com",
    "SENDER_PASSWORD": "your_app_password_here",  # Gmail App Password
    "RECIPIENTS":      ["recipient@example.com"],
    "SMTP_SERVER":     "smtp.gmail.com",
    "SMTP_PORT":       587,
    "SEND_TIME":       "08:00",   # Daily send time (24h format)
    "DATA_CSV":        "sales_data.csv",
}
```

### 🔑 Gmail App Password Setup
Gmail requires an **App Password** (not your login password):
1. Go to **Google Account → Security → 2-Step Verification → App Passwords**
2. Create a new App Password for "Mail"
3. Paste it into `SENDER_PASSWORD` in CONFIG

---

## 🚀 How to Run

### 1. Install dependencies
```bash
pip install pandas openpyxl schedule
```

### 2. Generate sample data
```bash
python generate_data.py
```

### 3. Test without sending (recommended first)
```bash
python email_bot.py --dry-run
```

### 4. Send one report immediately
```bash
python email_bot.py --now
```

### 5. Start the daily scheduler
```bash
python email_bot.py
```
The bot runs in the background and sends at the configured time every day.  
Press `Ctrl+C` to stop.

---

## 📊 What the Email Contains

**Email body** — Professional HTML email with a KPI summary table:
- Report date
- Total and completed orders
- Revenue (30-day, 7-day, yesterday)
- Top product and top sales rep
- Return rate

**Attachment** — A 4-sheet Excel report:

| Sheet | Contents |
|-------|----------|
| Summary | KPI cards for the current period |
| Daily Trend | Day-by-day revenue with bar chart |
| By Product | Revenue and units per product |
| Rep Leaderboard | Sales reps ranked by revenue with medals |

---

## 📋 Sample Log Output

```
2024-03-27 08:00:01  [INFO]  ═══════════════════════════════════════════
2024-03-27 08:00:01  [INFO]    SCHEDULED EMAIL REPORT BOT — JOB STARTED
2024-03-27 08:00:01  [INFO]  ═══════════════════════════════════════════
2024-03-27 08:00:01  [INFO]  📝  Building report...
2024-03-27 08:00:02  [INFO]  ✅  Report built → Daily_Sales_Report_20240327.xlsx
2024-03-27 08:00:02  [INFO]  📤  Sending email...
2024-03-27 08:00:04  [INFO]  ✅  Email sent successfully to: ['manager@company.com']
```

---

## 💡 Key Design Decisions

| Decision | Reason |
|----------|--------|
| `--dry-run` mode | Test the full pipeline safely without sending real emails |
| HTML email body | Professional appearance with KPI table visible in email preview |
| Separate `report_builder.py` | Report logic is reusable independently of the email system |
| Log file | Provides audit trail of every run — critical for production bots |
| App Password (not login password) | Follows Gmail security best practices |
| Temp file cleanup after send | Keeps the working directory clean automatically |

---

## 👤 Author

**Jatin Prasad**  
MS Robotics and Autonomous Systems (Systems Engineering)  
Ira A. Fulton Schools of Engineering, Arizona State University

*Part of a data analytics portfolio demonstrating automation, scheduling, and reporting skills.*
