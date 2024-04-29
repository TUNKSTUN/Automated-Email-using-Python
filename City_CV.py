import pandas as pd
import os

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

def main():
    # Define cities
    cities = [
        "Jubail", "Dammam", "Al Khobar", "Jeddah", "Riyadh",
        "Dhahran", "Al-Hassa", "Saihat", "Medinah", "Makkah", "Abha", "Rabigh"
    ]

    # Create subfolders for each city in the 'email_templates' directory
    for city in cities:
        city_folder = os.path.join('email_templates', city)
        os.makedirs(city_folder, exist_ok=True)

    # Load email template from Template.txt
    template = read_template('Template.txt')
    if template is None:
        return

    # Iterate through each city and its corresponding Excel file
    for city in cities:
        excel_file = os.path.join('cities', city, f'{city}.xlsx')
        try:
            df = pd.read_excel(excel_file)
        except FileNotFoundError:
            print(f"Error: Excel file '{excel_file}' not found.")
            continue

        # Iterate through each row in the DataFrame and generate individualized email templates
        for index, row in df.iterrows():
            name, company, email = row['Full Name'], row['Company Name'], row['Email']
            email_template = generate_email_template(name, company, template)
            
            # Create file path using email address as file name
            file_path = os.path.join('email_templates', city, f'{email}.txt')
            
            # Write email template to a separate text file named after the email address
            with open(file_path, 'w') as email_file:
                email_file.write(email_template)
            
if __name__ == "__main__":
    main()
