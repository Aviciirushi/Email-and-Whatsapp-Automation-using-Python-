import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import os
import re
import random
import time

# -----------------------
# 1. Configuration
# -----------------------
# 👉 Replace with your own email + app password before use
EMAIL_ADDRESS = "your_email@gmail.com"
EMAIL_PASSWORD = "your_app_password"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EXCEL_FILE = "sent_emails_list.xlsx"
MAX_EMAILS_PER_DAY = 100  # Safe Gmail daily limit (can be adjusted)

# -----------------------
# 2. Email Templates (Demo / Sample)
# -----------------------
def get_email_template(day):
    """
    Returns (subject, body) for a given scheduled day.
    These are demo messages — replace with your own content.
    """
    subject, body = "", ""

    if day == 1:
        subject = "Welcome to Our Mailing List"
        body = """Dear User,

Thank you for connecting with us! This is your first update.

We’ll be sharing:
✅ Industry insights
✅ Product updates
✅ Helpful resources

Stay tuned for more!

Best regards,
The Demo Team
"""

    elif day == 4:
        subject = "Case Study: How Businesses Grow With Us"
        body = """Hello,

Here’s a quick case study showing how companies achieved growth 
by adopting modern solutions.

👉 Would you like a personalized walkthrough?

Cheers,
The Demo Team
"""

    elif day == 10:
        subject = "Exclusive Benefits of Our Services"
        body = """Hi there,

What makes us different:
✅ Proven reliability
✅ High customer satisfaction
✅ Flexible solutions

👉 Should we prepare a custom plan for your needs?

Regards,
The Demo Team
"""

    elif day == 20:
        subject = "Industry Trends You Shouldn’t Miss"
        body = """Dear User,

We’re seeing strong demand for innovative solutions in multiple regions.

To help you stay ahead, we provide:
✅ Expert insights
✅ Reliable updates
✅ Actionable recommendations

👉 Want a tailored report for your industry?

Best,
The Demo Team
"""

    elif day == 40:
        subject = "Special Offer for Early Adopters"
        body = """Hello,

We’re closing this month’s schedule and wanted to check if you’d 
like to reserve a slot.

Early confirmation gives you:
✅ Priority access
✅ Exclusive discounts
✅ Dedicated support

👉 Should we reserve your spot?

Thanks,
The Demo Team
"""

    elif day == 90:
        subject = "Would You Like to Keep Receiving Updates?"
        body = """Hi,

If now isn’t the right time, no problem.  
We can still keep you updated on:
* New product launches
* Market insights
* Special offers

👉 Should we continue sending updates?

Thank you,
The Demo Team
"""

    return subject, body

# -----------------------
# 3. Send Email
# -----------------------
def send_email(to_addresses, subject, body):
    try:
        msg = MIMEMultipart()
        msg["From"] = EMAIL_ADDRESS
        msg["To"] = ", ".join(to_addresses)   # All in "To"
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        recipients = to_addresses

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, recipients, msg.as_string())
            print(f"📤 Email sent to {', '.join(to_addresses)} | Subject: {subject}")
            return True
    except Exception as e:
        print(f"❌ Failed to send to {', '.join(to_addresses)}: {e}")
        return False

# -----------------------
# 4. Main Logic — Sequential Schedule
# -----------------------
def run_email_schedule():
    if not os.path.exists(EXCEL_FILE):
        print("❌ Excel file not found.")
        return

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active

    headers = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]
    required_columns = [
        "Saved Date & Time", "To", "Email 1 Sent", "Email 2 Sent",
        "Email 3 Sent", "Email 4 Sent", "Email 5 Sent", "Email 6 Sent"
    ]

    for col in required_columns:
        if col not in headers:
            ws.cell(row=1, column=len(headers) + 1, value=col)
            headers.append(col)

    # Format: (index, column, threshold_days, template_day)
    email_schedule = [
        (1, "Email 1 Sent", 1, 1),
        (2, "Email 2 Sent", 4, 4),
        (3, "Email 3 Sent", 10, 10),
        (4, "Email 4 Sent", 20, 20),
        (5, "Email 5 Sent", 40, 40),
        (6, "Email 6 Sent", 90, 90)
    ]

    email_pattern = re.compile(r"[^@]+@[^@]+\.[^@]+")
    emails_sent_today = 0

    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        if emails_sent_today >= MAX_EMAILS_PER_DAY:
            print(f"⏹ Daily limit of {MAX_EMAILS_PER_DAY} emails reached.")
            break

        try:
            to_field = row[headers.index("To")].value
            date_value = row[headers.index("Saved Date & Time")].value

            if not to_field:
                continue

            emails = [e.strip() for e in str(to_field).split(",") if e.strip()]
            valid_emails = [e for e in emails if email_pattern.match(e)]

            if not valid_emails:
                continue

            to_emails = valid_emails

            if isinstance(date_value, datetime):
                date_added = date_value
            else:
                try:
                    date_added = datetime.strptime(str(date_value), "%Y-%m-%d %H:%M:%S")
                except:
                    date_added = datetime.strptime(str(date_value), "%d-%m-%Y %H:%M")

            now = datetime.now()
            days_since = (now - date_added).days

            for idx, col_name, threshold, template_day in email_schedule:
                col_index = headers.index(col_name)
                sent_status = row[col_index].value

                # Only send if previous was sent
                if idx > 1:
                    prev_status = row[headers.index(f"Email {idx-1} Sent")].value
                    if not prev_status or str(prev_status).strip().lower() != "sent":
                        break

                if sent_status and str(sent_status).strip().lower() == "sent":
                    continue

                if days_since >= threshold:
                    subject, body = get_email_template(template_day)
                    success = send_email(to_emails, subject, body)
                    if success:
                        ws.cell(row=i, column=col_index + 1, value="Sent")
                        emails_sent_today += 1

                        # Randomized Delay (12–18 minutes between emails)
                        if emails_sent_today < MAX_EMAILS_PER_DAY:
                            delay = random.randint(720, 1080)
                            print(f"⏳ Waiting {delay//60} minutes before next email...")
                            time.sleep(delay)
                    break

        except Exception as e:
            print(f"❌ Error processing row {i}: {e}")
            continue

    wb.save(EXCEL_FILE)
    print(f"✅ Email automation run complete. {emails_sent_today} emails sent today.")

# -----------------------
# 5. Trigger
# -----------------------
if __name__ == "__main__":
    run_email_schedule()
