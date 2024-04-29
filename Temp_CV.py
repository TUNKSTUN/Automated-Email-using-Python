import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

def extract_company_name(text):
    start_index = text.find("at ") + len("at ")
    end_index = text.find("that", start_index)
    if end_index == -1:
        end_index = text.find("that", start_index)
    company_name = text[start_index:end_index].strip()
    return company_name

def send_email(sender_name, sender_email, sender_password, receiver_email, subject, message, attachment_path):
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
    msg['From'] = f"{sender_name} <{sender_email}>"
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
    # Define cities
    cities = [
        "Jubail", "Dammam", "Al Khobar", "Jeddah", "Riyadh",
        "Dhahran", "Al-Hassa", "Saihat", "Medinah", "Makkah", "Abha", "Rabigh", "Hyderabad"
    ]
    
    # Get user input for city selection
    print("Available cities:")
    for idx, city in enumerate(cities, 1):
        print(f"{idx}. {city}")
    city_index = int(input("Select a city (enter the number): "))
    selected_city = cities[city_index - 1]

    # Set up email parameters
    attachment_path = 'Yahya.pdf'  # Path to your CV file
    sender_name = 'Yahya'
    # Get sender email from the name of the text files in the city directory
    sender_email = 'ykinwork1@gmail.com'
    sender_password = 'gwvo ossf ivpy pbmd'
    # Set email subject template
    subject_template = 'Application: Network Engineer Position at {}'

    # Read message from each text file in the selected city directory
    city_folder = f'email_templates/{selected_city}'
    for file_name in os.listdir(city_folder):
        if file_name.endswith('.txt'):
            receiver_email = file_name[:-4]  # Remove ".txt" extension
            message_file = os.path.join(city_folder, file_name)
            with open(message_file, 'r') as file:
                message = file.read()

            # Extract company name from message
            company_name = extract_company_name(message)

            # Format subject with extracted company name
            subject = subject_template.format(company_name)

            # Send email to each recipient
            send_email(sender_name, sender_email, sender_password, receiver_email, subject, message, attachment_path)

if __name__ == "__main__":
    main()
