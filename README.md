🤖 Automated WhatsApp & Email Lead Sender

Tired of manual follow-ups? This project automates WhatsApp messages and emails to your leads — hands-free!
Built with Python, Selenium, and Pandas, it fetches leads from Excel/IndiaMART exports, schedules personalized drip campaigns, and even integrates with Windows Task Scheduler for daily auto-runs.

✨ Features

📂 Import leads directly from Excel / IndiaMART exports

💬 Send automated WhatsApp messages using Selenium

📧 Deliver email follow-ups via SMTP

📅 Hands-free scheduling with Windows Task Scheduler

🔄 Run multi-step drip campaigns with multiple templates

📝 Logging system to track all messages sent

⚙️ Requirements

🐍 Python 3.9+

🌐 Google Chrome + ChromeDriver

📊 Excel file with leads (Name, Contact, Email, etc.)

Install dependencies:

pip install pandas selenium openpyxl

🚀 Quick Start
1️⃣ Clone the repo
git clone https://github.com/yourusername/whatsapp-email-automation.git
cd whatsapp-email-automation

2️⃣ Configure leads

Place your leads file (leads.xlsx) inside the project folder

Update message templates in messages.txt

3️⃣ Run manually
python main.py

4️⃣ Automate with Task Scheduler

Create a .bat file with:

python C:\path\to\main.py


Open Task Scheduler → Create Task → Set Trigger (Daily @ 9 AM) → Select .bat file

📊 Example Lead File (leads.xlsx)
Name	Phone Number	Email
John Doe	9876543210	john@email.com

Priya Patel	9123456780	priya@email.com
⚡ Roadmap

🎨 Add a GUI for easier usage

🔑 Support multiple WhatsApp accounts

📧 Integrate Gmail API for better email delivery

📊 Build an analytics dashboard

🛡️ Disclaimer

This project is for educational purposes only.
Please ensure compliance with WhatsApp & Email policies before use.

🔥 With this tool, you’ll never miss a follow-up again — automate your outreach & scale effortlessly!
