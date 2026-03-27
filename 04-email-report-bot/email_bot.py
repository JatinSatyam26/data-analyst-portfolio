"""
email_bot.py
────────────────────────────────────────────────────────────────
Scheduled Email Report Bot
Author : Jatin Prasad
Purpose: Automatically builds a daily sales Excel report and
         emails it as an attachment at a scheduled time every day.

Features:
  - Builds a fresh formatted Excel report from latest CSV data
  - Sends it via email with a professional HTML email body
  - Runs on a daily schedule (configurable time)
  - Logs every run with timestamp and status
  - Supports Gmail (App Password) or any SMTP server
  - Dry-run mode — test without sending a real email

Usage:
  python email_bot.py                  ← runs scheduler (sends daily at configured time)
  python email_bot.py --now            ← send report immediately (one shot)
  python email_bot.py --dry-run        ← build report + preview, no email sent

Configuration:
  Edit the CONFIG section below before running.
────────────────────────────────────────────────────────────────
"""

import smtplib
import os
import sys
import schedule
import time
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from report_builder import build_report

# ══════════════════════════════════════════════════════════════
#  ⚙️  CONFIGURATION — Edit these before running
# ══════════════════════════════════════════════════════════════
CONFIG = {
    # ── Sender (Gmail recommended) ─────────────────────────────
    "SENDER_EMAIL":    "your_email@gmail.com",
    "SENDER_PASSWORD": "your_app_password_here",   # Gmail App Password (not your login password)

    # ── Recipients ─────────────────────────────────────────────
    "RECIPIENTS": [
        "recipient1@example.com",
        # "recipient2@example.com",   # add more as needed
    ],

    # ── SMTP Server ────────────────────────────────────────────
    "SMTP_SERVER": "smtp.gmail.com",
    "SMTP_PORT":   587,

    # ── Schedule ───────────────────────────────────────────────
    "SEND_TIME": "08:00",          # 24-hour format — bot sends report at this time daily

    # ── Report Data Source ─────────────────────────────────────
    "DATA_CSV": "sales_data.csv",

    # ── Logging ────────────────────────────────────────────────
    "LOG_FILE": "bot_log.txt",
}
# ══════════════════════════════════════════════════════════════

# ── Logging Setup ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  [%(levelname)s]  %(message)s",
    handlers=[
        logging.FileHandler(CONFIG["LOG_FILE"]),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.getLogger(__name__)

# ── Email HTML Body ───────────────────────────────────────────
def build_email_body(kpis, report_date):
    rows_html = ""
    for k, v in kpis.items():
        rows_html += f"""
        <tr>
          <td style="padding:8px 16px;border-bottom:1px solid #eee;color:#555;font-weight:600;">{k}</td>
          <td style="padding:8px 16px;border-bottom:1px solid #eee;color:#E67E22;font-weight:700;">{v}</td>
        </tr>"""

    return f"""
    <html><body style="font-family:Calibri,Arial,sans-serif;background:#f7f7f7;margin:0;padding:0;">
      <div style="max-width:620px;margin:30px auto;background:#fff;border-radius:8px;
                  box-shadow:0 2px 8px rgba(0,0,0,0.1);overflow:hidden;">

        <!-- Header -->
        <div style="background:#145A32;padding:28px 32px;">
          <h1 style="color:#fff;margin:0;font-size:22px;">📊 Daily Sales Report</h1>
          <p style="color:#A9DFBF;margin:6px 0 0;font-size:13px;">{report_date}</p>
        </div>

        <!-- Body -->
        <div style="padding:28px 32px;">
          <p style="color:#444;font-size:14px;margin-top:0;">
            Your automated daily sales report is attached. Here is today's summary:
          </p>

          <table style="width:100%;border-collapse:collapse;margin-top:16px;">
            <thead>
              <tr style="background:#1E8449;">
                <th style="padding:10px 16px;color:#fff;text-align:left;">Metric</th>
                <th style="padding:10px 16px;color:#fff;text-align:left;">Value</th>
              </tr>
            </thead>
            <tbody>
              {rows_html}
            </tbody>
          </table>

          <p style="color:#888;font-size:12px;margin-top:24px;">
            📎 The full Excel report (4 sheets: Summary, Daily Trend, By Product, Rep Leaderboard)
            is attached to this email.<br><br>
            This report was generated and sent automatically by the
            <strong>Scheduled Email Report Bot</strong>.
          </p>
        </div>

        <!-- Footer -->
        <div style="background:#f0f0f0;padding:16px 32px;text-align:center;">
          <p style="color:#aaa;font-size:11px;margin:0;">
            Auto-generated | Data Analyst Portfolio — Jatin Prasad | ASU Fulton Schools of Engineering
          </p>
        </div>

      </div>
    </body></html>
    """

# ── Send Email ────────────────────────────────────────────────
def send_email(report_path, kpis, dry_run=False):
    today       = datetime.today().strftime("%B %d, %Y")
    subject     = f"📊 Daily Sales Report — {today}"
    body_html   = build_email_body(kpis, today)

    msg = MIMEMultipart("alternative")
    msg["From"]    = CONFIG["SENDER_EMAIL"]
    msg["To"]      = ", ".join(CONFIG["RECIPIENTS"])
    msg["Subject"] = subject
    msg.attach(MIMEText(body_html, "html"))

    # Attach Excel report
    with open(report_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f'attachment; filename="{os.path.basename(report_path)}"'
    )
    msg.attach(part)

    if dry_run:
        log.info("🔍  DRY RUN — Email NOT sent. Preview below:")
        log.info(f"    To      : {CONFIG['RECIPIENTS']}")
        log.info(f"    Subject : {subject}")
        log.info(f"    Report  : {report_path}")
        log.info("    ✅  Everything looks good. Set dry_run=False to send for real.")
        return True

    try:
        log.info(f"📤  Connecting to {CONFIG['SMTP_SERVER']}:{CONFIG['SMTP_PORT']}...")
        with smtplib.SMTP(CONFIG["SMTP_SERVER"], CONFIG["SMTP_PORT"]) as server:
            server.ehlo()
            server.starttls()
            server.login(CONFIG["SENDER_EMAIL"], CONFIG["SENDER_PASSWORD"])
            server.sendmail(
                CONFIG["SENDER_EMAIL"],
                CONFIG["RECIPIENTS"],
                msg.as_string()
            )
        log.info(f"✅  Email sent successfully to: {CONFIG['RECIPIENTS']}")
        return True
    except smtplib.SMTPAuthenticationError:
        log.error("❌  Authentication failed. Check your email and App Password in CONFIG.")
        return False
    except smtplib.SMTPException as e:
        log.error(f"❌  SMTP error: {e}")
        return False
    except Exception as e:
        log.error(f"❌  Unexpected error: {e}")
        return False

# ── Main Job ──────────────────────────────────────────────────
def run_job(dry_run=False):
    log.info("=" * 55)
    log.info("  SCHEDULED EMAIL REPORT BOT — JOB STARTED")
    log.info("=" * 55)

    # Step 1: Build report
    log.info("📝  Building report...")
    try:
        report_path, kpis = build_report(CONFIG["DATA_CSV"])
    except Exception as e:
        log.error(f"❌  Report build failed: {e}")
        return

    # Step 2: Send email
    log.info("📤  Sending email...")
    success = send_email(report_path, kpis, dry_run=dry_run)

    # Step 3: Clean up report file
    if success and not dry_run:
        os.remove(report_path)
        log.info(f"🗑️   Temp report file cleaned up.")

    log.info("=" * 55)

# ── Entry Point ───────────────────────────────────────────────
if __name__ == "__main__":
    args = [a.lower() for a in sys.argv[1:]]

    if "--dry-run" in args:
        # Build report + preview, no email sent
        print("\n🔍  Running in DRY RUN mode — no email will be sent.\n")
        run_job(dry_run=True)

    elif "--now" in args:
        # Send immediately (one shot)
        print("\n🚀  Sending report NOW (one shot)...\n")
        run_job(dry_run=False)

    else:
        # Run on daily schedule
        send_time = CONFIG["SEND_TIME"]
        print(f"\n⏰  Scheduler started — report will be sent daily at {send_time}")
        print(f"    Data source : {CONFIG['DATA_CSV']}")
        print(f"    Recipients  : {CONFIG['RECIPIENTS']}")
        print(f"    Log file    : {CONFIG['LOG_FILE']}")
        print("    Press Ctrl+C to stop.\n")

        schedule.every().day.at(send_time).do(run_job)

        # Also run immediately on start so you can verify it works
        log.info(f"🕐  Next run scheduled at {send_time} daily.")
        while True:
            schedule.run_pending()
            time.sleep(30)
