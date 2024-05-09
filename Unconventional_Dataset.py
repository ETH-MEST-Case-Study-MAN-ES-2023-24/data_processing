import pandas as pd
import numpy as np

excel_file = 'Unconventional_Dataset.xlsx'
input_sheet_name = 'Unconventional_Dataset'
output_sheet_name = 'Definitiv'
df = pd.read_excel(excel_file, sheet_name=input_sheet_name)
df_copied = df.copy()

# Values to filter for in the "classification" column
filter_values = ['Metro stations', 'Wastewater treatment plants', 'Food retail', 'Food production']

# Filter rows where the "classification" column contains any of the specified values
df_filtered = df_copied[df_copied['Category'].astype(str).str.contains('|'.join(filter_values), case=False, na=False)]
# print(len(df_filtered))

# Define conditions and values for the new columns
conditions = [
    df_filtered['Category'] == 'Metro stations',
    df_filtered['Category'] == 'Wastewater treatment plants',
    df_filtered['Category'] == 'Food retail',
    df_filtered['Category'] == 'Food production'
]

# Define values corresponding to each condition
T_min = [5, 8, 40, 20]  # Adjust these values as needed
T_max = [35, 15, 70, 40]  # Adjust these values as needed

# Use numpy.select to assign values based on conditions
df_filtered['Tsource_min [°C]'] = np.select(conditions, T_min, default=np.nan)
df_filtered['Tsource_max [°C]'] = np.select(conditions, T_max, default=np.nan)

column_to_export = ['Latitude', 'Longitude', 'FacilityName', 'Category', 'Tsource_min [°C]', 'Tsource_max [°C]', 'QL_TJ']
# Export the filtered DataFrame to a new sheet
with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    # Export the original DataFrame 
    df.to_excel(writer, sheet_name=input_sheet_name, index=False)

    df_filtered[column_to_export].rename(columns={'FacilityName': 'Name', 'Classification': 'Sector name', 'QL_TJ': 'Energy [TJ]'}).to_excel(writer, sheet_name=output_sheet_name, index=False)
