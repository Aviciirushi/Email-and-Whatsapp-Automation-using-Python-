import imaplib
import email
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os

# Gmail credentials
EMAIL_USER = "digimedexxim@gmail.com"
EMAIL_PASS = "bbdt pwfs hqok wflp"

# Excel file path
excel_file = "sent_emails_list.xlsx"

# Load existing workbook or create new
if os.path.exists(excel_file):
    wb = load_workbook(excel_file)
    ws = wb.active
    # Ensure header exists
    if ws.max_row == 0 or ws.cell(row=1, column=1).value != "Saved Date & Time":
        ws.append([
            "Saved Date & Time", "To",
            "Email 1 Sent", "Email 2 Sent", "Email 3 Sent",
            "Email 4 Sent", "Email 5 Sent"
        ])
    # Load existing recipients (normalized)
    saved_recipients = {
        str(row[1].value).strip().lower()
        for row in ws.iter_rows(min_row=2) if row[1].value
    }
else:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sent Emails"
    ws.append([
        "Saved Date & Time", "To",
        "Email 1 Sent", "Email 2 Sent", "Email 3 Sent",
        "Email 4 Sent", "Email 5 Sent"
    ])
    saved_recipients = set()

# Connect to Gmail IMAP
mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login(EMAIL_USER, EMAIL_PASS)
mail.select('"[Gmail]/Sent Mail"')  # Access Sent Mail folder

# Search for all sent emails
status, messages = mail.search(None, "ALL")
email_ids = messages[0].split()

print(f"📧 Found {len(email_ids)} sent emails in Gmail.")

saved_count = 0

for eid in email_ids:
    status, msg_data = mail.fetch(eid, "(RFC822)")
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])

            # Get recipient(s) exactly as Gmail stores
            to_ = msg.get("To", "Unknown").strip()

            # Normalize for uniqueness check
            to_normalized = to_.lower().strip()

            # Save only if not already saved
            if to_normalized not in saved_recipients:
                current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                ws.append([current_datetime, to_, "", "", "", "", ""])
                saved_recipients.add(to_normalized)
                saved_count += 1
                print(f"✅ Saved new entry: {to_}")

# Save to Excel
wb.save(excel_file)
print(f"💾 Saved {saved_count} new unique leads to {excel_file}")

# Logout
mail.logout()
