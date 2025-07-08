import gspread
from google.oauth2.service_account import Credentials
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
from email.utils import formataddr
from googleapiclient.discovery import build
from datetime import datetime, timedelta
import time
import imaplib

# --- CONFIGURATION ---
SHEET_NAME = "Expo-Sales-Management"
SHEET_TAB = "OB-speakers"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
CREDS_FILE = "/etc/secrets/service_account.json"
creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
gc = gspread.authorize(creds)
sheet = gc.open(SHEET_NAME).worksheet(SHEET_TAB)
sheet_id = sheet.spreadsheet.id

SMTP_SERVER = "mail.b2bgrowthexpo.com"
SMTP_PORT = 587
EMAIL_SENDER = "speakersengagement@b2bgrowthexpo.com"
EMAIL_PASSWORD = "jH!Ra[9q[f68"

# --- COLORS ---
RED = {"red": 1.0, "green": 0.0, "blue": 0.0}
NEON_SKY_BLUE = {"red": 0.0, "green": 1.0, "blue": 1.0}
YELLOW_RGB = {"red": 1.0, "green": 1.0, "blue": 0.0}

# --- EMAIL SIGNATURE ---
HTML_SIGNATURE = """
<p>If you would like to schedule a meeting with me,<br>
please use the link below:<br>
<a href="https://tidycal.com/nagendra/b2b-discovery-call" target="_blank">https://tidycal.com/nagendra/b2b-discovery-call</a></p>
<p style="margin-top: 30px;">
Thanks & Regards,<br>
<strong>Nagendra Mishra</strong><br>
Director | B2B Growth Hub<br>
Mo: +44 7913 027482<br>
Email: <a href="mailto:nagendra@b2bgrowthexpo.com">nagendra@b2bgrowthexpo.com</a><br>
<a href="https://www.b2bgrowthexpo.com" target="_blank">www.b2bgrowthexpo.com</a>
</p>
<p style="font-size: 13px; color: #888;">
If you don’t want to hear from me again, please let me know.
</p>
"""

# --- FOLLOW-UP EMAIL TEMPLATES ---
FOLLOWUP_BODIES = [
    """<p>Thank you for submitting your interest to participate as a Speaker at the <strong>{show}</strong> B2B Growth Expo.</p>
<p>To proceed with your application, we request you to register using the URL below:<br>
<a href="https://b2bgrowthexpo.com/speakers-registration/">https://b2bgrowthexpo.com/speakers-registration/</a></p>
<p>Please confirm once you have registered, and feel free to ask if you need any clarification.</p>""",

    """<p>Dear {name},</p>
<p>This is to follow up on your registration for the Speaking opportunity.<br>
Did you manage to register? If not, do you need any more information?</p>""",

    """<p>Dear {name},</p>
<p>I understand that sometimes, interest is expressed out of curiosity; however, it is not a burning desire.</p>
<p>If that is the case with you, then we can understand why you still haven't registered.</p>
<p>I do not want to flood your inbox with unnecessary messages. Can you please confirm if you are still looking to pursue this?</p>""",

    """<p>Dear {name},</p>
<p>Without sounding too rude, we are eager to onboard you, but given the very time-sensitive nature of the business, we do not want to waste our time on unnecessary follow-ups.</p>
<p>I request you to kindly save my time and respond in a simple YES / NO to cut to the chase.</p>"""
]

# --- UTILITY FUNCTIONS ---
def is_yellow(color):
    return round(color.get("red", 0), 1) == 1.0 and round(color.get("green", 0), 1) == 1.0 and round(color.get("blue", 0), 1) == 0.0

def get_row_color(service, row_index):
    result = service.spreadsheets().get(
        spreadsheetId=sheet_id,
        ranges=[f"{SHEET_TAB}!A{row_index + 1}"],
        fields="sheets.data.rowData.values.userEnteredFormat.backgroundColor"
    ).execute()
    try:
        return result['sheets'][0]['data'][0]['rowData'][0]['values'][0]['userEnteredFormat']['backgroundColor']
    except:
        return {}

def update_color(service, row_index, color):
    service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body={
        "requests": [{
            "repeatCell": {
                "range": {
                    "sheetId": sheet._properties['sheetId'],
                    "startRowIndex": row_index,
                    "endRowIndex": row_index + 1
                },
                "cell": {"userEnteredFormat": {"backgroundColor": color}},
                "fields": "userEnteredFormat.backgroundColor"
            }
        }]
    }).execute()

def is_24hrs_passed(timestamp_str):
    try:
        return (datetime.now() - datetime.strptime(timestamp_str, "%Y-%m-%d %H:%M:%S")) >= timedelta(hours=24)
    except:
        return True

def send_followup_email(to_email, subject, html_body):
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = formataddr(("Nagendra Mishra", EMAIL_SENDER))
        msg["To"] = to_email
        msg.attach(MIMEText(html_body, "html"))
        raw_msg = msg.as_string()

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)

        imap = imaplib.IMAP4_SSL(SMTP_SERVER)
        imap.login(EMAIL_SENDER, EMAIL_PASSWORD)
        imap.append("INBOX.Sent", "", imaplib.Time2Internaldate(time.time()), raw_msg.encode("utf8"))
        imap.logout()
        print(f"✅ Sent: {to_email}")
        return True
    except Exception as e:
        print(f"❌ Failed to send email to {to_email}: {e}")
        return False

# --- MAIN PROCESSING ---
def process_batches():
    all_rows = sheet.get_all_values()
    if len(all_rows) < 2:
        print("⚠️ No data.")
        return

    header = all_rows[0]
    data = all_rows[1:]
    service = build("sheets", "v4", credentials=creds)

    status_idx = header.index("Status")
    ts_idx = header.index("Follow-up Timestamp") if "Follow-up Timestamp" in header else len(header)

    if "Follow-up Timestamp" not in header:
        sheet.update_cell(1, ts_idx + 1, "Follow-up Timestamp")

    for i, row in enumerate(data, start=2):
        padded = row + [""] * (len(header) + 2)
        name = padded[header.index("First_Name")].strip()
        email = padded[header.index("Email")].strip()
        show = padded[header.index("Show")].strip()
        response = padded[header.index("Email-Response")].strip().lower()
        status = padded[status_idx].strip()
        last_ts = padded[ts_idx].strip()

        # Handle "Action Required"
        if response == "action required" and status.lower() != "action required":
            current_color = get_row_color(service, i - 1)
            if current_color != {"red": 0.29, "green": 0.53, "blue": 0.91}:
                update_color(service, i - 1, {"red": 0.29, "green": 0.53, "blue": 0.91})  # Blue
            sheet.update_cell(i, status_idx + 1, "Action Required")
            continue

        # Handle "Offer Rejected"
        if response == "offer rejected" and status.lower() != "offer rejected":
            current_color = get_row_color(service, i - 1)
            if current_color != RED:
                update_color(service, i - 1, RED)
            sheet.update_cell(i, status_idx + 1, "Offer Rejected")
            continue

        # Skip if no email or not interested
        if not email or response != "interested":
            continue

        color = get_row_color(service, i - 1)
        if "email sent -1" in status.lower() and is_yellow(color):
            print(f"🟡 Row {i} turned yellow. Halting.")
            continue

        # Determine follow-up index
        followup_index = -1
        if status.lower().startswith("email sent -"):
            try:
                followup_index = int(status.lower().split("-")[-1])
            except:
                followup_index = -1
        if followup_index == -1:
            followup_index = 0
        if followup_index >= 4 or not is_24hrs_passed(last_ts):
            continue

        html_body = FOLLOWUP_BODIES[followup_index].format(name=name, show=show) + HTML_SIGNATURE
        subject = f"Speaker Follow-Up – {show}"

        if send_followup_email(email, subject, html_body):
            new_status = f"Email Sent -{followup_index + 1}" if followup_index < 3 else "All Followups Done"
            sheet.update_cell(i, status_idx + 1, new_status)
            sheet.update_cell(i, ts_idx + 1, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

            if followup_index == 0:
                if is_yellow(color):
                    update_color(service, i - 1, {})  # Remove yellow
                update_color(service, i - 1, NEON_SKY_BLUE)
            elif followup_index == 3:
                update_color(service, i - 1, RED)

while True:
    print("🔁 Running automation...")
    try:
        process_batches()
        print("✅ Run complete. Sleeping 2 hours...\n")
    except Exception as e:
        print(f"❌ Automation error: {e}")
    time.sleep(7200)
