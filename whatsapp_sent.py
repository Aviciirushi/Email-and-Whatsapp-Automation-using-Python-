import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import urllib.parse
import time
import random

# -----------------------------
# 1. File Paths
# -----------------------------
# 👉 Change to your Excel leads file + ChromeDriver path
EXCEL_PATH = "path/to/indiamart_leads.xlsx"
CHROMEDRIVER_PATH = "path/to/chromedriver.exe"

# -----------------------------
# 2. Load Excel Leads
# -----------------------------
df = pd.read_excel(EXCEL_PATH)

# -----------------------------
# 3. Message Templates
# -----------------------------
messages = {
    1: """Hello Sir/Mam, I hope you’re doing well.  
It was great connecting with you earlier about our pharmaceutical exports. At *Digimed Exxim*, we specialize in delivering high-quality, reliable products that meet international standards, ensuring your sourcing process is smooth and hassle-free.  

To better understand your requirements, could you share which product categories are your current priority?  
""",

    2: """Hi Sir/Mam,  

Just checking in again. Many of our clients in the USA, UK, France and Australia rely on us for a consistent supply of Weight Loss products, Steroids, and Tapentadol, while markets like **France, Singapore, and Malaysia** prefer our Ayurvedic & Herbal range.  

All our distributors in USA/Canada have reduced delays by 35% in 3 months after switching to our supply chain.  

I’d be happy to explore how we can help your business achieve similar results. Would you be open to a quick call this week to discuss specifics?  
""",

    3: """Hello Sir/Mam,  

Thank you for your time and consideration over the past weeks regarding our pharmaceutical export solutions. As a final note, I’d like to extend an exclusive introductory offer to support your business—that’s a special discount and a complimentary consultation tailored to your needs.  

If now isn’t the perfect time, no worries—just let me know if you’d prefer to reconnect in the future or have any questions. I’m always here to assist and ensure you get the best value and service with *Digimed Exxim*.  

Looking forward to your response and wishing you ongoing success!  
"""
}

# -----------------------------
# 4. Setup Selenium
# -----------------------------
driver = webdriver.Chrome(service=Service(CHROMEDRIVER_PATH))
driver.get("https://web.whatsapp.com/")

print("📱 Please scan the QR code in WhatsApp Web...")
input("✅ Press Enter once you are logged in...")

# -----------------------------
# 5. Helper Functions
# -----------------------------
def is_whatsapp_number(phone_number: str) -> bool:
    """
    Check if the given phone number is registered on WhatsApp.
    """
    driver.get(f"https://web.whatsapp.com/send?phone={phone_number}&text&app_absent=0")
    try:
        WebDriverWait(driver, 6).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true" and @data-tab]'))
        )
        return True
    except:
        return False

def send_message(phone_number: str, message: str) -> bool:
    """
    Send a WhatsApp message to the given phone number.
    """
    encoded_message = urllib.parse.quote(message)
    driver.get(f"https://web.whatsapp.com/send?phone={phone_number}&text={encoded_message}")

    try:
        send_button = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Send"]'))
        )
        send_button.click()
        time.sleep(2)
        return True
    except Exception as e:
        print(f"❌ Failed to send to {phone_number}: {e}")
        return False

# -----------------------------
# 6. Process Leads (max 20/day)
# -----------------------------
sent_count = 0
today = datetime.today()

for idx, row in df.iterrows():
    if sent_count >= 20:
        break

    phone = str(row["Phone"])
    if pd.isna(phone):
        continue

    # Clean number format
    phone = phone.replace("+", "").replace("-", "").replace(" ", "")
    if len(phone) < 8:
        continue

    # Skip permanently if previously marked Skipped
    if any(str(row.get(col)) == "Skipped" for col in ["WhatsApp 1 Sent", "WhatsApp 2 Sent", "WhatsApp 3 Sent"]):
        print(f"🚫 {phone} permanently marked as Skipped, ignoring...")
        continue

    # Safe date parsing
    lead_date = pd.to_datetime(row["Date"], dayfirst=True, errors="coerce")
    if pd.isna(lead_date):
        continue
    days_passed = (today - lead_date).days

    # Decide which message to send
    msg_number = None
    if pd.isna(row.get("WhatsApp 1 Sent")) and days_passed >= 1:
        msg_number = 1
    elif pd.isna(row.get("WhatsApp 2 Sent")) and days_passed >= 5:
        msg_number = 2
    elif pd.isna(row.get("WhatsApp 3 Sent")) and days_passed >= 12:
        msg_number = 3

    if msg_number:
        print(f"➡️ Checking WhatsApp availability for {phone}...")

        if not is_whatsapp_number(phone):
            print(f"🚫 {phone} not on WhatsApp, marking permanently as Skipped...")
            df.at[idx, f"WhatsApp {msg_number} Sent"] = "Skipped"
            continue

        # If valid, send message
        print(f"📲 Sending WhatsApp {msg_number} to {phone}...")
        if send_message(phone, messages[msg_number]):
            df.at[idx, f"WhatsApp {msg_number} Sent"] = today.strftime("%Y-%m-%d")
            print(f"✅ WhatsApp {msg_number} sent to {phone}")
            sent_count += 1

            # Delay between messages to avoid spam
            if sent_count < 20:
                wait_time = random.randint(1800, 7200)  # 30–120 minutes
                print(f"⏳ Waiting {wait_time//60} min before next lead...")
                time.sleep(wait_time)

# -----------------------------
# 7. Save Excel
# -----------------------------
df.to_excel(EXCEL_PATH, index=False)
print("💾 Excel updated with sent/skipped status!")

driver.quit()
