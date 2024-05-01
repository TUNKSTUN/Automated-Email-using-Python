import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas as pd

def extract_company_name(html_text):
    start_index = html_text.find("at <strong>")
    if start_index == -1:
        return None
    end_index = html_text.find("</strong>.", start_index + len("at <strong>"))
    if end_index == -1:
        return None
    company_name = html_text[start_index + len("at <strong>"):end_index].strip()
    return company_name



def generate_email_template(name, company, template):
    # Replace placeholders in the template with actual data
    email_template = template.replace('[name]', name).replace('[Company]', company)
    return email_template

def read_template(file_path):
    try:
        with open(file_path, 'r') as file:
            template = file.read()
        return template
    except FileNotFoundError:
        print(f"Error: Template file '{file_path}' not found.")
        return None

def send_email(sender_name, sender_email, sender_password, receiver_email, subject, message, attachment_path, badges):
    # Set up the SMTP server
    smtp_server = 'smtp.gmail.com'
    port = 587  # For starttls
    smtp_username = sender_email
    smtp_password = sender_password

    try:
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
        message_text = MIMEText(message, 'html')
        msg.attach(message_text)

        # Add badges in HTML format
        badges_html = "<html><body><p>---</p><h4>"
        for badge in badges:
            badges_html += f"<img src='{badge}' alt='Badge' style='width:300px;height:auto;'>"
        badges_html += "<h4></body></html>"
        msg.attach(MIMEText(badges_html, 'html'))

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

        # Print success log
        print(f"Email sent successfully to {receiver_email}.")

    except Exception as e:
        # Print failure log
        print(f"Failed to send email to {receiver_email}: {str(e)}")

def create_dummy_email_template():
    template = read_template('Template.txt')

    if template is None:
        return

    # Define the city and email address
    city = "Hyderabad"
    dummy_email_address = "johnwick4learning@gmail.com"
    name = "Fahad"
    company = "Disney Land"

    email_template = generate_email_template(name, company, template)

    # Construct the directory path
    directory = os.path.join("email_templates", city)

    # Construct the file path
    file_path = os.path.join(directory, f"{dummy_email_address}.txt")

    # Check if the directory exists, and create it if it doesn't
    if not os.path.exists(directory):
        os.makedirs(directory)

    # Check if the file already exists, and delete it
    if os.path.isfile(file_path):
        print(f"File Exists!")
        return file_path

    # Write the content to the file
    with open(file_path, "w") as file:
        file.write(email_template)
        print(f"Dummy email template created: {file_path}")

def main():
    create_dummy_email_template()
    # Load email template from Template.txt
    template = read_template('Template.txt')
    if template is None:
        return
    # Define cities
    cities = [
        "Jubail", "Dammam", "Al Khobar", "Jeddah", "Riyadh",
        "Dhahran", "Al-Hassa", "Saihat", "Medinah", "Makkah", "Abha", "Rabigh", "Hyderabad"
    ]

    # Create subfolders for each city in the 'email_templates' directory
    for city in cities:
        city_folder = os.path.join('email_templates', city)
        os.makedirs(city_folder, exist_ok=True)

    

    # Load the Excel file
    df = pd.read_excel('Companies_KSA.xls')

    # Create the 'cities' directory if it doesn't exist
    cities_dir = 'cities'
    os.makedirs(cities_dir, exist_ok=True)

    # Iterate through each city name
    for city in cities:
        # Filter rows based on the city name
        filtered_df = df[df['City'].str.contains(city, case=False)]
        if not filtered_df.empty:
            # Create subfolder for the city within the 'cities' directory
            city_dir = os.path.join(cities_dir, city)
            os.makedirs(city_dir, exist_ok=True)

            # Save the filtered DataFrame to a new Excel file within the city's subfolder
            file_name = f'{city}.xlsx'
            file_path = os.path.join(city_dir, file_name)
            filtered_df.to_excel(file_path, index=False)

            # Iterate through each row in the DataFrame and generate individualized email templates
            for index, row in filtered_df.iterrows():
                name, company, email = row['Full Name'], row['Company Name'], row['Email']
                email_template = generate_email_template(name, company, template)
                
                # Create file path using email address as file name
                email_file_path = os.path.join('email_templates', city, f'{email}.txt')
                
                # Write email template to a separate text file named after the email address
                with open(email_file_path, 'w') as email_file:
                    email_file.write(email_template)

    # Get user input for city selection
    print("Available cities:")
    for idx, city in enumerate(cities, 1):
        print(f"{idx}. {city}")
    city_index = int(input("Select a city (enter the number): "))
    selected_city = cities[city_index - 1]

        # Set up email parameters
    attachment_path = 'Yahya.pdf'  # Path to your CV file
    sender_name = 'Yahya'
    sender_email = os.environ.get('SENDER_EMAIL')
    sender_password = os.environ.get('SENDER_PASSWORD')
    subject_template = 'Application: Network Engineer Position at {}'
    badges = ['https://github.com/TUNKSTUN/Automated-Email-using-Python/blob/main/BADGES.png?raw=true']

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
            send_email(sender_name, sender_email, sender_password, receiver_email, subject, message, attachment_path, badges)


if __name__ == "__main__":
    main()