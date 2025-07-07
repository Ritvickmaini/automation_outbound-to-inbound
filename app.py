import gspread
from google.oauth2.service_account import Credentials
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
from email.utils import formataddr
from googleapiclient.discovery import build
import time 

# --- CONFIGURATION ---
SHEET_NAME = "Expo-Sales-Management"
SHEET_TAB = "OB-speakers"
STATUS_COL_NAME = "Status"  # Add this column in the Sheet
ROW_COLOR_ACTION_REQUIRED = {"red": 0.29, "green": 0.53, "blue": 0.91}  # Blue
ROW_COLOR_OFFER_REJECTED = {"red": 1.0, "green": 0.0, "blue": 0.0}

# --- GOOGLE SHEETS SETUP ---
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
CREDS_FILE = "/etc/secrets/service_account.json"
creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
gc = gspread.authorize(creds)
sheet = gc.open(SHEET_NAME).worksheet(SHEET_TAB)
sheet_id = sheet.spreadsheet.id

# --- EMAIL CONFIG ---
SMTP_SERVER = "mail.b2bgrowthexpo.com"
SMTP_PORT = 587
EMAIL_SENDER = "speakersengagement@b2bgrowthexpo.com"
EMAIL_PASSWORD = "jH!Ra[9q[f68"

# --- HTML TEMPLATE ---
HTML_TEMPLATE = """
<html>
  <body style="font-family: Arial, sans-serif; font-size: 15px; color: #333; background-color: #ffffff; padding: 20px;">
    <p>Dear {name},</p>
    <p>
      I hope this message finds you well.<br><br>
      Thank you for showing interest in speaking at our upcoming <strong>{show}</strong>.
      This exciting event will bring together industry leaders, innovators, and professionals 
      for a day of connection, collaboration, and the exchange of valuable insights.
      We would be honoured to welcome you as one of our speakers.
    </p>
    <p>
      While this is an unpaid opportunity, speaking at the Expo offers several key benefits:
    </p>
    <ul>
      <li>Increased visibility and recognition within your industry</li>
      <li>Opportunities to expand your professional network</li>
      <li>A platform to showcase your expertise to a diverse and engaged audience</li>
    </ul>
    <p>
      Our previous events have drawn a dynamic mix of participants, including startup founders, 
      SME owners, corporate executives, and other influential figures from across various sectors—
      ensuring a high-quality audience for your session.
    </p>
    <p>
      If you are interested, please let us know your availability at your earliest convenience 
      so we can reserve your speaking slot and discuss any specific needs you may have.
    </p>
    <p>
      Thank you for considering this invitation. I look forward to the possibility of working with you 
      and hope to welcome you as a valued speaker at the {show}.
    </p>
    <p>
      If you would like to schedule a meeting with me,<br>
      please use the link below:<br>
      <a href="https://tidycal.com/nagendra/b2b-discovery-call" target="_blank">https://tidycal.com/nagendra/b2b-discovery-call</a>
    </p>
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
  </body>
</html>
"""

# --- FUNCTIONS ---
def send_email(to_email, name, show):
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = f"Invitation to Speak at {show}"
        msg["From"] = formataddr(("Nagendra Mishra", EMAIL_SENDER))
        msg["To"] = to_email

        html_content = HTML_TEMPLATE.format(name=name, show=show)
        msg.attach(MIMEText(html_content, "html"))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
            print(f"✅ Email sent to {name} at {to_email}")
    except Exception as e:
        print(f"❌ Failed to send email to {to_email}: {e}")

def color_row(row_index, color):
    service = build("sheets", "v4", credentials=creds)
    body = {
        "requests": [{
            "repeatCell": {
                "range": {
                    "sheetId": sheet._properties['sheetId'],
                    "startRowIndex": row_index,
                    "endRowIndex": row_index + 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": color
                    }
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        }]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=body).execute()
    print(f"🎨 Colored row {row_index + 1}")

# --- PROCESS EACH ROW ---
rows = sheet.get_all_records()
def get_column_index(column_name):
    header_row = sheet.row_values(1)
    return header_row.index(column_name) + 1  # 1-based index
def process_sheet():
    rows = sheet.get_all_records()
for idx, row in enumerate(rows, start=2):  # start=2 since header is row 1
    response = row.get("Email-Response", "").strip().lower()
    email = row.get("Email", "").strip()
    name = row.get("First_Name", "").strip()
    show = row.get("Show", "").strip()
    status = row.get("Status", "").strip().lower()

    print(f"[{idx}] Status: {response} | Email: {email} | Name: {name} | Show: {show}")

    if response == "interested" and email and status != "email sent":
        send_email(email, name, show)
        sheet.update_cell(idx, get_column_index("Status"), "Email Sent")
    elif response == "action required":
        color_row(idx - 1, ROW_COLOR_ACTION_REQUIRED)
    elif response == "offer rejected":
        color_row(idx - 1, ROW_COLOR_OFFER_REJECTED)
while True:
    print("⏳ Starting automation run...")
    try:
        process_sheet()
        print("✅ Run complete. Sleeping for 2 hours...\n")
    except Exception as e:
        print(f"❌ Error during run: {e}")
    time.sleep(7200)  # Sleep for 2 hours