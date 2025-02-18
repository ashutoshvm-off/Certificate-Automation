import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os

# Configuration
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAIL_ADDRESS = 'xxxxxxxx@gmail.com'#add your mail id here
EMAIL_PASSWORD = ''#add your app password here
EXCEL_FILE = 'data.xlsx'
DOCUMENTS_FOLDER = 'out'  # Folder where your documents are stored

# Read names and email addresses from Excel
df = pd.read_excel(EXCEL_FILE)

# Function to send email
def send_email(to_address, name):
    subject = f"xxxxx" #change the content
    body = f"Dear {name},\n\n xxxxxx\n\nBest regards,\nxxx\n xxxx"

    # Create the email
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_address
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    #Attach the document if it exists
    document_path = os.path.join(DOCUMENTS_FOLDER, f"{name}.pdf")  # Change extension if needed
    if os.path.exists(document_path):
       with open(document_path, 'rb') as file:
           attach = MIMEApplication(file.read(), _subtype='pdf')
           attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(document_path))
           msg.attach(attach)
    return msg
    # Send the email
    # with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
    #     # server.starttls()
    #     # server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    #     server.send_message(msg)
    #     

# Loop through each row in the DataFrame
# for index, row in df.iterrows():
#     name = row['name']
#     email_address = row['Email Address']
#     send_email(email_address, name)


if __name__== '__main__':
     with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)


        for index, row in df.iterrows():
            name = row['name']
            email_address = row['Email Address']
            msg = send_email(email_address, name)
            server.send_message(msg)
            print(f"Email sent to {name} at {email_address}")
