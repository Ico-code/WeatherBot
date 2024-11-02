import os
import pandas
from openpyxl import load_workbook
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv

from weather_bot import log_email_send

errorLogFilePath = "error_log.xlsx";

def logErrorToExcel(error_details={"errMsg":"api key is empty","errLVL":"medium", "errLocation":"tasks.py"}):
    """
    Write error message in error_log.xlsx
    Parameters:
        error_details should contain all the data that is needed for logging the info into excel.
        This would include the following:
            errMsg:
                message containing a description of the error
            errLVL:
                how critical is the error
            errLocation:
                where, or what file did the error occur in

    # Example usage
    error_info = {
        "Error Level": "Error",
        "Location": "data_processing.py, line 54",
        "Error Message": "File not found",
    }
    logErrorToExcel(error_info)
    """
    try:
        workbook = load_workbook(errorLogFilePath)
        sheet = workbook.active
    except FileNotFoundError:
        from openpyxl import Workbook
        workbook = Workbook()
        sheet = workbook.active
        # Write headers if file is new
        sheet.append(["Timestamp", "Error Level", "Location","Error Message"])
    # Append error details as a new row
    currentTime = datetime.now().strftime("%Y-%m-%d %H:%M:%S");
    sheet.append([
        currentTime,
        error_details.get("errLVL", ""),
        error_details.get("errLocation",""),
        error_details.get("errMsg", ""),
    ])

    # Save the workbook
    workbook.save(errorLogFilePath)

    errorContent = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            table {{
                font-family: Arial, sans-serif;
                border-collapse: collapse;
                width: 100%;
            }}
            th, td {{
                border: 1px solid #dddddd;
                text-align: left;
                padding: 8px;
            }}
            th {{
                background-color: #f2f2f2;
            }}
        </style>
    </head>
    <body>

    <h2>Error Log</h2>

    <table>
        <tr>
            <th>Timestamp</th>
            <td>{currentTime}</td>
        </tr>
        <tr>
            <th>Error Level</th>
            <td>{error_details.get("errLVL", "")}</td>
        </tr>
        <tr>
            <th>Location</th>
            <td>{error_details.get("errLocation", "")}</td>
        </tr>
        <tr>
            <th>Error Message</th>
            <td>{error_details.get("errMsg", "")}</td>
        </tr>
    </table>

    </body>
    </html>
    """
    notifyAdministrator(errorContent)

def notifyAdministrator(body):
    print("Notifying...")

    recipient_emails = getAdministrators();

    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    sender_email = os.getenv("EMAIL_ADDRESS")
    sender_password = os.getenv("EMAIL_PASSWORD")

    # Compose the email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(recipient_emails)
    msg['Subject'] = "There has been an error, while making Weather Report"
    msg.attach(MIMEText(body, 'html'))
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Secure the connection
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_emails, msg.as_string())
        print("Email sent successfully.")
        log_email_send("Success", body)  # Log success if email is sent
    except Exception as e:
        print(f"Failed to send email: {e}")
        log_email_send("Failed", body)  # Log failure if email is not sent
    
def getAdministrators():
    """
    Gets a list of all administrator emails
    """
    print("Getting administrators...")

    df = pandas.read_excel("Users.xlsx")
    
    # Check the columns (for reference)
    print("Columns in the file:", df.columns)

    # Filter rows where the Role is "administrator"
    admins = df[df["Role"].str.lower() == "administrator"]
    
    # Extract the email addresses
    admin_emails = admins["User List"].tolist()
    
    return admin_emails