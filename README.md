Automated WhatsApp & Email Lead Sender

This project automates sending WhatsApp messages and emails to leads using Python, Selenium, and Pandas. It fetches leads from Excel/IndiaMART, schedules personalized drip messages, and automates daily follow-ups. With Windows Task Scheduler integration, everything runs hands-free.

✨ Features

📂 Fetch leads directly from Excel files / IndiaMART exports

💬 Send automated WhatsApp messages with Selenium

📧 Automate email follow-ups with SMTP

📅 Schedule daily tasks using Windows Task Scheduler

🔄 Supports multiple message templates & drip campaigns

📝 Logs all sent messages for tracking

⚙️ Requirements

Python 3.9+

Google Chrome + ChromeDriver

Excel file with leads (Name, Contact, Email, etc.)

Install dependencies:

pip install pandas selenium openpyxl

🚀 Usage

Clone the repo

git clone https://github.com/yourusername/whatsapp-email-automation.git
cd whatsapp-email-automation


Configure leads

Place your leads file (leads.xlsx) in the project folder

Update message templates in messages.txt

Run script manually

python main.py


Automate with Task Scheduler

Create a .bat file:

python C:\path\to\main.py


Open Task Scheduler → Create Task → Set trigger (Daily at 9 AM) → Select .bat

📊 Example Lead File (leads.xlsx)
Name	Phone Number	Email
John Doe	9876543210	john@email.com

Priya Patel	9123456780	priya@email.com
⚡ Roadmap

 Add GUI for easier use

 Support multiple WhatsApp accounts

 Integrate Gmail API for better email delivery

 Add analytics dashboard

🛡️ Disclaimer

This project is for educational purposes only. Use responsibly and ensure compliance with WhatsApp/Email policies.
