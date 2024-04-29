import pandas as pd
import os

# Define the list of city names
city_names = [
    'Jubail', 'Dammam', 'Al Khobar', 'Jeddah', 'Riyadh',
    'Dhahran', 'Al-Hassa', 'Saihat', 'Medinah', 'Makkah', 'Abha', 'Rabigh'
]

# Load the Excel file
df = pd.read_excel('Companies_KSA.xls')

# Create the 'cities' directory if it doesn't exist
cities_dir = 'cities'
os.makedirs(cities_dir, exist_ok=True)

# Iterate through each city name
for city in city_names:
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
