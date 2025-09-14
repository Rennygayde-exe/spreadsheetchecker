import os
import json
import requests
import smtplib
from email.mime.text import MIMEText
from dotenv import load_dotenv
load_dotenv()

EMAIL = os.getenv("EMAIL_USER")
PASSWORD = os.getenv("EMAIL_PASS")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
RECIPIENTS = os.getenv("RECIPIENTS", EMAIL).split(",")

STATE_FILE = "last_stats.json"
API_URL = "https://script.google.com/macros/s/AKfycbyIBV7uMsR2rIMpQffdSIU3y7Ddb6zwIEesd3YUYUQfE_1srS_20dEC-6QTWxeglFdi/exec"

def get_stats():
    resp = requests.get(API_URL)
    resp.raise_for_status()
    return resp.json()

def load_last():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r") as f:
            return json.load(f)
    return None

def save_current(data):
    with open(STATE_FILE, "w") as f:
        json.dump(data, f)

def send_email_update(stats):
    body = f"""Spreadsheet stats updated:

Year:  {stats.get('year')}
Month: {stats.get('month')}
Week:  {stats.get('week')}
"""
    msg = MIMEText(body)
    msg["Subject"] = "Spreadsheet Stats Updated"
    msg["From"] = EMAIL
    msg["To"] = ", ".join(RECIPIENTS)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL, PASSWORD)
        server.send_message(msg)

def daily_check():
    current = get_stats()
    last = load_last()

    if not last or current != last:
        print("Margot did some spreadsheets! Sending Update letter")
        send_email_update(current)
        save_current(current)
    else:
        print("Margot hasnt done any spreadsheets today, check back tomarrow.")

if __name__ == "__main__":
    daily_check()
