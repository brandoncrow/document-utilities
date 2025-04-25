"""
file_copy.py

Copies and renames files based on an Excel mapping.

Usage:
    python file_copy.py <path_to_file.xlsx> <new_directory> <original_file_path_column> <new_filename_column>

Example:
    python file_copy.py "mapping.xlsx" "D:/renamed_docs" "OriginalPath" "NewFileName"
"""

import os
import sys
import shutil
import pandas as pd

def main():
    # check for the correct number of command-line arguments.
    if len(sys.argv) != 5:
        print("Usage: python file_copy.py <path_to_file.xlsx> <new_directory> <original_file_path_column> <new_filename_column>")
        sys.exit(1)
    
    input_excel = sys.argv[1]
    new_dir = sys.argv[2]
    orig_col = sys.argv[3]
    new_filename_col = sys.argv[4]
    
    # verify the input Excel file exists.
    if not os.path.isfile(input_excel):
        print(f"Error: {input_excel} is not a valid file.")
        sys.exit(1)
    
    # create the new directory if it doesn't exist.
    if not os.path.exists(new_dir):
        os.makedirs(new_dir)
    
    # read the Excel file.
    print("Reading Excel file...")
    df = pd.read_excel(input_excel)
    
    # check required columns.
    if orig_col not in df.columns or new_filename_col not in df.columns:
        print(f"Error: The Excel file must contain columns '{orig_col}' and '{new_filename_col}'.")
        print("Available columns:", list(df.columns))
        sys.exit(1)
    
    new_paths = []  # store new file paths.
    error_log = []  # store error messages.
    
    total_rows = len(df)
    print(f"Processing {total_rows} rows...")
    
    for index, row in df.iterrows():
        original_path = row[orig_col]
        new_name = str(row[new_filename_col])
        
        if not os.path.isfile(original_path):
            error_message = f"File not found: {original_path} (Row {index + 2})"
            print("Warning:", error_message)
            error_log.append(error_message)
            new_paths.append(None)
            continue
        
        new_file_path = os.path.join(new_dir, new_name)
        
        try:
            shutil.copy2(original_path, new_file_path)
        except Exception as e:
            error_message = f"Error copying file {original_path} to {new_file_path}: {e} (Row {index + 2})"
            print("Error:", error_message)
            error_log.append(error_message)
            new_paths.append(None)
            continue
        
        new_paths.append(new_file_path)
        
        # print progress every 10 rows.
        if (index + 1) % 100 == 0 or (index + 1) == total_rows:
            print(f"Processed {index + 1}/{total_rows} rows.")
    
    df["New File Path"] = new_paths
    
    base_name = os.path.basename(input_excel)
    name_no_ext, _ = os.path.splitext(base_name)
    output_excel = os.path.join(new_dir, f"{name_no_ext}_updated.xlsx")
    df.to_excel(output_excel, index=False)
    
    if error_log:
        error_log_file = os.path.join(new_dir, "error_log.txt")
        with open(error_log_file, "w") as log_file:
            for error in error_log:
                log_file.write(error + "\n")
        print(f"Some errors occurred. See the error log at: {error_log_file}")
    
    print(f"Processing complete. Updated Excel file saved to: {output_excel}")

if __name__ == "__main__":
    main()