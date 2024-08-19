import pandas as pd
import os
import re

# Directory containing the Excel files
directory = '/content/'  # Replace with your directory path

# Initialize an empty DataFrame to hold the consolidated data
consolidated_data = pd.DataFrame()

# Loop through each file in the directory
for filename in os.listdir(directory):
    if filename.endswith(".xlsx"):  # Ensure it's an Excel file
        # Load the workbook
        workbook_path = os.path.join(directory, filename)
        xls = pd.ExcelFile(workbook_path)

        # Loop through each sheet in the workbook
        for sheet_name in xls.sheet_names:
            if sheet_name not in ('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'November', 'December'):
                continue
            # Read the sheet into a DataFrame, but only select columns A to P
            df = pd.read_excel(workbook_path, sheet_name=sheet_name, usecols="A:P", skiprows= 1, header= None)

            # Add a new column for the sheet name
            df['Sheet Name'] = sheet_name
            df['File Name'] = re.sub(r'\d.*', '', filename)

            # Append the data to the consolidated DataFrame
            consolidated_data = pd.concat([consolidated_data, df], ignore_index=True)

# Save the consolidated data to a new Excel file
consolidated_data.to_excel('Consolidated_Data.xlsx', index=False, header=False)
