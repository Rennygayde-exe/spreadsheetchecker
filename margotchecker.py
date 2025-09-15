import os
import json
import smtplib
import requests
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

load_dotenv()

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
RECIPIENTS = os.getenv("RECIPIENTS", "").split(",")

STATS_URL = "https://script.google.com/macros/s/AKfycbyIBV7uMsR2rIMpQffdSIU3y7Ddb6zwIEesd3YUYUQfE_1srS_20dEC-6QTWxeglFdi/exec"
LAST_STATS_FILE = "last_stats.json"

def fetch_stats():
    resp = requests.get(STATS_URL, timeout=10)
    resp.raise_for_status()
    return resp.json()

def load_last_stats():
    if not os.path.exists(LAST_STATS_FILE):
        return None
    with open(LAST_STATS_FILE, "r") as f:
        return json.load(f)

def save_last_stats(stats):
    with open(LAST_STATS_FILE, "w") as f:
        json.dump(stats, f, indent=2)
    print("last_stats.json updated")

def send_email_update(stats, diffs):
    plain_body = f""" Margots Daily Spreadsheet stats:

Year:  {stats.get('year')}  ({diffs['year']:+})
Month: {stats.get('month')}  ({diffs['month']:+})
Week:  {stats.get('week')}  ({diffs['week']:+})
"""

    html_body = f"""
    <html>
      <body style="font-family: Arial, sans-serif; color: #333;">
        <h2 style="color:#2c3e50;">Spreadsheet Stats Updated</h2>
        <p><b>Year:</b> {stats.get('year')} <span style="color: #27ae60;">({diffs['year']:+})</span></p>
        <p><b>Month:</b> {stats.get('month')} <span style="color: #27ae60;">({diffs['month']:+})</span></p>
        <p><b>Week:</b> {stats.get('week')} <span style="color: #27ae60;">({diffs['week']:+})</span></p>
        <hr style="margin:20px 0;">
        <p style="font-size: 0.9em; color: #777;">
          This update was sent automatically by Margots Spreadsheet Tracker
        </p>
      </body>
    </html>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "Margot did a new Spreadsheet!"
    msg["From"] = EMAIL
    msg["To"] = ", ".join(RECIPIENTS)

    msg.attach(MIMEText(plain_body, "plain"))
    msg.attach(MIMEText(html_body, "html"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL, PASSWORD)
        server.send_message(msg)

def daily_check():
    current = get_stats()
    last = load_last()

    if last:
        diffs = {
            "year": current["year"] - last.get("year", 0),
            "month": current["month"] - last.get("month", 0),
            "week": current["week"] - last.get("week", 0),
        }
    else:
        diffs = {"year": 0, "month": 0, "week": 0}

    if not last or current != last:
        print("Margot did some spreadsheets! Sending Update letter")
        send_email_update(current, diffs)
        save_current(current)
    else:
        print("Margot hasnt done any spreadsheets today, check back tomarrow.")
if __name__ == "__main__":
    main()
