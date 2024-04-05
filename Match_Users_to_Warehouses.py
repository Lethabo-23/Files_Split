import os
import pandas as pd

# Load the Excel file
excel_file = 'Users and Warehouses.xlsx'
users_df = pd.read_excel(excel_file, sheet_name='Users')
wh_df = pd.read_excel(excel_file, sheet_name='WH_Site_Location')

# Perform cartesian product
cartesian_product = pd.merge(users_df.assign(key=1), wh_df.assign(key=1), on='key').drop('key', axis=1)

# Specify the output folder
output_folder = 'Users_WH_Matched'

# Create the output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)
else:
    # Clear the existing files in the folder
    files = os.listdir(output_folder)
    for file in files:
        os.remove(os.path.join(output_folder, file))

# Get the number of output files from the user
num_output_files = int(input("Enter the number of output files: "))

# Calculate the chunk size to split the data into equal parts
chunk_size = len(cartesian_product) // num_output_files

# Save the result to multiple Excel files with the specified names in the output folder
for i in range(num_output_files):
    start_index = i * chunk_size
    end_index = (i + 1) * chunk_size if i < num_output_files - 1 else len(cartesian_product)

    subset_result = cartesian_product.iloc[start_index:end_index]
    output_excel = os.path.join(output_folder, f'User {i + 1} - Users and WH Matched.xlsx')
    subset_result.to_excel(output_excel, index=False, sheet_name='Sheet1')
    print(f"Matching completed. Result saved to {output_excel}")

print("All matching completed.")
