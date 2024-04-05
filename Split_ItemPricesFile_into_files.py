import os
import glob
import pandas as pd
import shutil

def prepare_output_directory(output_directory):
    # Check if the output directory exists
    if os.path.exists(output_directory):
        # Clear the contents of the existing directory
        for file in os.listdir(output_directory):
            file_path = os.path.join(output_directory, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
    else:
        # Create the output directory if it doesn't exist
        os.makedirs(output_directory)

def split_excel(input_directory, output_directory, num_files):
    # Check if the input directory exists
    if not os.path.exists(input_directory):
        print("Input directory does not exist.")
        return

    # Prepare the output directory
    prepare_output_directory(output_directory)

    # Find the specific input Excel file
    input_file = os.path.join(input_directory, 'Items.xlsx')

    # Read the input Excel file
    df = pd.read_excel(input_file)

    # Calculate the number of rows per output file
    rows_per_file = len(df) // num_files

    for i in range(num_files):
        # Calculate the starting and ending indices for each output file
        start = i * rows_per_file
        end = start + rows_per_file if i < num_files - 1 else None

        # Create the output Excel file name with the desired format
        output_file_name = f"User {i + 1} - Activate Item Prices.xlsx"

        # Write the selected rows to the output Excel file with the specified sheet name
        df_subset = df.iloc[start:end]
        output_path = os.path.join(output_directory, output_file_name)
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_subset.to_excel(writer, index=False, sheet_name='Sheet1')

if __name__ == "__main__":
    input_directory = os.getcwd()  # Use the current working directory as the input directory
    output_directory = os.path.join(input_directory, "Split_items_files_to_EA")  # Specify the output directory within the current directory

    # Split the Excel file
    num_files = int(input("Enter the number of files to split into: "))
    split_excel(input_directory, output_directory, num_files)

    print(f"Excel file 'Items.xlsx' in the root folder split into {num_files} files in '{output_directory}'.")
