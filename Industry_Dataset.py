import numpy as np
import pandas as pd

# Import excel
excel_file = 'Industry_Dataset.xlsx'
input_sheet_name = 'D5_1_Industry_Dataset'
df = pd.read_excel(excel_file, sheet_name=input_sheet_name)

# Temperature levels
T_level_1 = 25.0
T_level_2 = 55.0

# Create new columns
df['Energy [TJ]'] = np.nan
df['Temperature [°C]'] = np.nan



# Coefficients for the linear system
A = np.array([[1, -T_level_1], [1, -T_level_2]])

# Iterate over each row of df with iterrows
for index, row in df.iterrows():
    # Extract the values for the right-hand side vector (b) (or Q)
    b = row[['level_1_Tj', 'level_2_Tj']].astype('float64').values

    # Solve the linear system A*x = b
    x = np.linalg.solve(A, b)

    # Store the optimization results in the DataFrame
    df.at[index, 'Energy [TJ]'] = x[0]
    df.at[index, 'Temperature [°C]'] = x[0] / x[1]

# Export excel
output_sheet_name = 'Definitiv'
columns_to_export = ['Latitude', 'Longitude', 'CompanyNam', 'Eurostat_N']

# Create a new ExcelWriter object
with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    # Export the original DataFrame to the 'D5_1_Industry_Dataset' sheet
    df.to_excel(writer, sheet_name=input_sheet_name, index=False)

    # Export selected columns with changed names to the 'Definitiv' sheet
    df_selected = df[['Latitude', 'Longitude', 'CompanyNam', 'Eurostat_N']].rename(columns={'CompanyNam': 'Name', 'Eurostat_N': 'Sector name'})
    df_selected.to_excel(writer, sheet_name=output_sheet_name, index=False)
    df[['Temperature [°C]', 'Energy [TJ]']].to_excel(writer, sheet_name=output_sheet_name, index=False, startcol=4)
