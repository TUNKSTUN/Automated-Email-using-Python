import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd

def send_email(sender_email, sender_password, receiver_email, subject, message, attachment_path):
    # Set up the SMTP server
    smtp_server = 'smtp.gmail.com'
    port = 587  # For starttls
    smtp_username = sender_email
    smtp_password = sender_password

    # Create a secure connection to the SMTP server
    server = smtplib.SMTP(smtp_server, port)
    server.starttls()
    server.login(smtp_username, smtp_password)

    # Create a multipart message and set headers
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    # Add message body
    msg.attach(MIMEText(message, 'plain'))

    # Open the file to be sent
    with open(attachment_path, "rb") as attachment:
        # Add the file as application/octet-stream
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {attachment_path}",
    )

    # Add attachment to message
    msg.attach(part)

    # Send the email
    server.send_message(msg)

    # Close the connection
    server.quit()

def main():
    # Load data from Excel sheet containing recipients' email addresses
    df = pd.read_excel('test.xlsx')

    # Set up email parameters
    sender_email = 'johnwick4learning@gmail.com'
    sender_password = 'sdvj nydy txsn ayoi'
    subject = 'Test'
    message = 'Test'
    attachment_path = 'BLANK.pdf'  # Path to your CV file

    # Send emails to each recipient in the DataFrame
    for index, row in df.iterrows():
        receiver_email = row['Email']
        send_email(sender_email, sender_password, receiver_email, subject, message, attachment_path)

if __name__ == "__main__":
    main()
