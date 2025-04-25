"""
copy_docs_to_subfolder.py

This script copies documents from a source folder into structured subfolders under a destination directory.
The destination path is constructed from metadata in an Excel file (EntityType and Number).
Any missing or failed files are logged.

Typical Use Case: Organizing legal/land documents after a bulk download using metadata.
"""

import os
import shutil
import pandas as pd

# configuration
source_folder = r'C:\Path\To\LocalDocs\DownloadFolder'
destination_folder = r'C:\Path\To\Destination\Documents'
excel_file_path = r'C:\Path\To\MappingFile\Documents.xlsx'

missing_files_log = r'C:\Path\To\Logs\missing_files.txt'
copy_errors_log = r'C:\Path\To\Logs\copy_errors.txt'

# Clear existing log files (if any)
with open(missing_files_log, "w") as f:
    f.write("Missing Files Log:\n")

with open(copy_errors_log, "w") as f:
    f.write("Copy Errors Log:\n")

# read the Excel file
try:
    df = pd.read_excel(excel_file_path)
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit(1)

# iterate over each row in the DataFrame
for index, row in df.iterrows():
    entity_type = str(row["EntityType"]).strip()
    entity_number = str(row["Number"]).strip()
    document_name = str(row["DocumentName"]).strip()

    # construct the source file path
    src_file_path = os.path.join(source_folder, document_name)

    # construct the destination directory and file path:
    dest_dir = os.path.join(destination_folder, entity_type, entity_number)
    dest_file_path = os.path.join(dest_dir, document_name)

    # create the destination directory if it doesn't already exist
    try:
        os.makedirs(dest_dir, exist_ok=True)
    except Exception as e:
        with open(copy_errors_log, "a") as error_file:
            error_file.write(f"Error creating directory {dest_dir}: {e}\n")
        continue

    # check if the source file exists; if not, log the missing file and continue
    if not os.path.isfile(src_file_path):
        with open(missing_files_log, "a") as log_file:
            log_file.write(f"{document_name} not found for entity {entity_number}.\n")
        continue

    # attempt to copy the file
    try:
        if not os.path.exists(dest_file_path):
            shutil.copy2(src_file_path, dest_file_path)
            print(f"Copied: {document_name} to {dest_dir}")
        else:
            print(f"Skipped: {document_name} already exists in {dest_dir}")
    except Exception as e:
        with open(copy_errors_log, "a") as error_file:
            error_file.write(f"Error copying {document_name} to {dest_dir}: {e}\n")
            