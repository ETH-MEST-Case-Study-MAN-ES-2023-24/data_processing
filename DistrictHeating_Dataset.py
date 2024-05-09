import numpy as np
import pandas as pd
from geopy.geocoders import Nominatim
from unidecode import unidecode
import pandas as pd

# Import excel 
excel_file = 'DistrictHeating_Dataset.xlsx' 
input_sheet_name = 'Cities_Dataset'
output_sheet_name = 'Definitiv'
df = pd.read_excel(excel_file, sheet_name=input_sheet_name)
df_copied = df.copy()

# Filter rows where a specific column contains 'DH'
df_filtered = df_copied[df_copied['Type'].astype(str).str.contains('DH', case=False, na=False)]

df_filtered['Tsource_min [°C]'] = 50
df_filtered['Tsource_max [°C]'] = 100
df_filtered['Energy [TJ]'] = df_filtered['HD_345_TJ']

# Initialize geocoding service
geolocator = Nominatim(user_agent="CaseStudyETH")

# Geocode each city and update Latitude and Longitude columns
for index, row in df_filtered.iterrows():
    city_name = str(row['Placename_'])
    region_name = str(row['NAME_ASCI'])
    
    # Clean city name from special characters
    cleaned_city_name = unidecode(city_name)

    # Perform geocoding only if cleaned_city_name is different from city_name
    if cleaned_city_name == city_name:
        location = geolocator.geocode(city_name, timeout=10)

        if location:
            df_filtered.at[index, 'Latitude'] = location.latitude
            df_filtered.at[index, 'Longitude'] = location.longitude
        else:
            df_filtered.at[index, 'Latitude'] = "not geolocalized"
            df_filtered.at[index, 'Longitude'] = "not geolocalized"
    else:
        location = geolocator.geocode(region_name, timeout=10)

        if location:
            df_filtered.at[index, 'Latitude'] = location.latitude
            df_filtered.at[index, 'Longitude'] = location.longitude
        else:
            df_filtered.at[index, 'Latitude'] = "not geolocalized"
            df_filtered.at[index, 'Longitude'] = "not geolocalized"

column_to_export = ['Latitude', 'Longitude', 'Placename_', 'Type', 'Tsource_min [°C]', 'Tsource_max [°C]', 'Power [MW]', 'Shape__Are', 'Shape__Len']

# Create a new ExcelWriter object
with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    # Export the original DataFrame 
    df.to_excel(writer, sheet_name=input_sheet_name, index=False)

    # Export selected columns with changed names to the 'Definitiv' sheet
    df_filtered[column_to_export].rename(columns={'Placename_': 'Name', 'Type': 'Sector name'}).to_excel(writer, sheet_name=output_sheet_name, index=False)
