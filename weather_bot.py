import os
import smtplib
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
import zipfile
import datetime

# Load environment variables
load_dotenv()

def send_weather_email(weather_data, recipient_emails):
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    sender_email = os.getenv("EMAIL_ADDRESS")
    sender_password = os.getenv("EMAIL_PASSWORD")

    # Compose the email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(recipient_emails)
    msg['Subject'] = "Weather Report"

    # Format the weather data into an HTML table
    body = "<h2>Weather Report</h2><table border='1'><tr><th>Parameter</th><th>Value</th></tr>"
    for key, value in weather_data.items():
        body += f"<tr><td>{key}</td><td>{value}</td></tr>"
    body += "</table><br>"

    msg.attach(MIMEText(body, 'html'))  # Attach the HTML body

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Secure the connection
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_emails, msg.as_string())
        print("Email sent successfully.")
        log_email_send("Success")  # Log success if email is sent
        return True
    except Exception as e:
        print(f"Failed to send email: {e}")
        log_email_send("Failed")  # Log failure if email is not sent
        return False

def log_email_send(status="Failed", log_file="Log.xlsx"):
    try:
        wb = load_workbook(log_file)  # Attempt to load existing workbook
        sheet = wb.active
    except (FileNotFoundError, zipfile.BadZipFile):  # Handle both file not found and bad zip file
        wb = Workbook()  # Create a new workbook
        sheet = wb.active
        sheet.append(["Timestamp", "Status"])  # Add headers

    # Append log entry
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sheet.append([timestamp, status])  # Log the email send status

    wb.save(log_file)  # Save workbook
    print("Logged email send to Log.xlsx")
