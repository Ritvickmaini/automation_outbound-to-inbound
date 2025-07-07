import gspread
from google.oauth2.service_account import Credentials
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
from email.utils import formataddr
from googleapiclient.discovery import build
import time 
import imaplib

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
        # Create the email
        msg = MIMEMultipart("alternative")
        msg["Subject"] = f"Invitation to Speak at {show}"
        msg["From"] = formataddr(("Nagendra Mishra", EMAIL_SENDER))
        msg["To"] = to_email

        html_content = HTML_TEMPLATE.format(name=name, show=show)
        msg.attach(MIMEText(html_content, "html"))

        # Convert message to full MIME string
        raw_message = msg.as_string()

        # --- SMTP: Send the email ---
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
            print(f"✅ Email sent to {name} at {to_email}")

        # --- IMAP: Save to "Sent" folder ---
        imap = imaplib.IMAP4_SSL("mail.b2bgrowthexpo.com")
        imap.login(EMAIL_SENDER, EMAIL_PASSWORD)
        imap.append("INBOX.Sent", "", imaplib.Time2Internaldate(time.time()), raw_message.encode("utf8"))
        imap.logout()
        print("📩 Email saved to Sent folder")
        return True

    except Exception as e:
        print(f"❌ Failed to send/save email to {to_email}: {e}")
        return False

    except Exception as e:
        print(f"❌ Failed to send/save email to {to_email}: {e}")

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

def get_col_letter(col_idx):
    result = ""
    while col_idx >= 0:
        result = chr(col_idx % 26 + 65) + result
        col_idx = col_idx // 26 - 1
    return result
  
def process_batches():
    all_rows = sheet.get_all_values()
    if len(all_rows) < 2:
        print("⚠️ Sheet is empty or only contains headers.")
        return

    header = all_rows[0]
    data_rows = all_rows[1:]
    status_col_index = header.index("Status")+1  # 0-based
    total_rows = len(data_rows)

    for batch_start in range(0, total_rows, 100):
        batch_end = min(batch_start + 100, total_rows)
        print(f"\n📦 Processing rows {batch_start + 2} to {batch_end + 1}")
        batch = data_rows[batch_start:batch_end]
        print(f"🔢 Batch size: {len(batch)}")

        status_updates = []
        format_requests = []

        for i, row in enumerate(batch, start=batch_start + 2):
            padded_row = (row + [""] * len(header))[:len(header)]
            row_data = dict(zip(header, padded_row))

            response = row_data.get("Email-Response", "").strip().lower()
            email = row_data.get("Email", "").strip()
            name = row_data.get("First_Name", "").strip()
            show = row_data.get("Show", "").strip()
            status = row_data.get("Status", "").strip().lower()

            print(f"[{i}] {response} | {email} | {name} | {show}")
            if not email:
                print(f"⚠️ Missing email at row {i}, skipping.")
                continue

            if response == "interested" and status != "email sent":
                if send_email(email, name, show):
                    status_updates.append((i, "Email Sent"))

            elif response == "action required" and status != "action required":
                format_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet._properties['sheetId'],
                            "startRowIndex": i - 1,
                            "endRowIndex": i
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": ROW_COLOR_ACTION_REQUIRED
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                })

            elif response == "offer rejected" and status != "offer rejected":
                format_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet._properties['sheetId'],
                            "startRowIndex": i - 1,
                            "endRowIndex": i
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": ROW_COLOR_OFFER_REJECTED
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                })

        if status_updates:
            col_letter = get_col_letter(status_col_index)
            data = [{
                "range": f"{SHEET_TAB}!{col_letter}{row_num}",
                "values": [[status_text]]
            } for row_num, status_text in status_updates]

            service = build("sheets", "v4", credentials=creds)
            service.spreadsheets().values().batchUpdate(
                spreadsheetId=sheet_id,
                body={"valueInputOption": "USER_ENTERED", "data": data}
            ).execute()
            print(f"📝 Updated {len(status_updates)} status cells")
            print(f"✅ Status column updated for rows: {[row_num for row_num, _ in status_updates]}")

        if format_requests:
            service = build("sheets", "v4", credentials=creds)
            service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id,
                body={"requests": format_requests}
            ).execute()
            print(f"🎨 Formatted {len(format_requests)} rows")

        time.sleep(2)  # Prevent rate limit

# --- Replace this in your while loop ---
while True:
    print("⏳ Starting automation run...")
    try:
        process_batches()
        print("✅ All batches processed. Sleeping for 2 hours...\n")
    except Exception as e:
        print(f"❌ Error during run: {e}")
    time.sleep(7200)

