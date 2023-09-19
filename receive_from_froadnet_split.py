import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Define the input folder path
input_folder = 'data/Receive_from_roadnet'

# Define the input file path
input_file = os.path.join(input_folder, 'sessionIDs.xlsx')

# Define the output folder path
output_folder = os.path.join(input_folder, 'inbound')

# Clear the 'inbound' folder if it exists
if os.path.exists(output_folder):
    for file in os.listdir(output_folder):
        file_path = os.path.join(output_folder, file)
        if os.path.isfile(file_path):
            os.remove(file_path)
else:
    os.makedirs(output_folder)

# Allow the user to select an option for the output file name
print("Select an option for the output file name:")
print("1. 12H")
print("2. 19H")

option = input("Enter the option (1 or 2): ")

# Validate user input and set the suffix accordingly
if option == '1':
    suffix = "12H"
elif option == '2':
    suffix = "19H"
else:
    print("Invalid option. Using default '12H'.")
    suffix = "12H"

# Load the Excel file
df = pd.read_excel(input_file)

# Create a dictionary to store dataframes for each unique Session ID
user_number = 1  # Initialize user number

for session_id, group in df.groupby("Session ID"):
    # Create a new workbook and add a worksheet
    workbook = Workbook()
    worksheet = workbook.active

    # Rename the worksheet to "Sheet 1"
    worksheet.title = "Sheet 1"

    # Set the "Session ID" as the header in the first row
    worksheet.append(["Session ID"])

    # Create a generator of rows from the DataFrame (excluding the header)
    rows = dataframe_to_rows(group[['Session ID']], index=False, header=False)

    # Populate the worksheet with rows from the DataFrame
    for row in rows:
        worksheet.append(row)

    # Save the workbook with the specified file name including the selected suffix
    output_file = os.path.join(output_folder, f'User {user_number} - {suffix} - {session_id} - Receive load from Roadnet into D365.xlsx')
    workbook.save(output_file)

    user_number += 1  # Increment user number

print("Excel sheets created successfully.")
