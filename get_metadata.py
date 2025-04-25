"""
get_metadata.py

Recursively collects metadata for all files in a given directory.

Each file's metadata includes:
- Full file path
- Filename
- File extension
- File size (in KB)
- File hash (MD5)

Output: An Excel file named `metadata.xlsx` in the scanned directory.

Usage:
    python get_metadata.py <directory_path>
"""

import os
import sys
import hashlib

def check_required_modules():
    missing_modules = []
    try:
        import pandas as pd
    except ModuleNotFoundError:
        missing_modules.append("pandas")
    try:
        import openpyxl
    except ModuleNotFoundError:
        missing_modules.append("openpyxl")
    if missing_modules:
        print("Error: The following required modules are missing: " + ", ".join(missing_modules))
        print("Please install them using pip, for example:")
        print("    pip install " + " ".join(missing_modules))
        sys.exit(1)

# check for required modules before doing any heavy processing.
check_required_modules()

import pandas as pd

def get_file_metadata(file_path):
    """
    Returns a dictionary with metadata for the given file:
      - file_path: full path including the filename
      - filename: just the file name with extension
      - file_type: the file extension (as seen in file explorer)
      - file_size: size in kilobytes (float rounded to 2 decimals)
      - file_hash: MD5 hash of the file's content
    """
    # compute MD5 hash using buffered reading
    md5_hash = hashlib.md5()
    try:
        with open(file_path, 'rb') as f:
            while chunk := f.read(4096):
                md5_hash.update(chunk)
        file_hash = md5_hash.hexdigest()
    except Exception:
        file_hash = None

    try:
        file_size = round(os.path.getsize(file_path) / 1024.0, 2)
    except Exception:
        file_size = None

    filename = os.path.basename(file_path)
    _, file_extension = os.path.splitext(filename)
    
    return {
        "file_path": file_path,
        "filename": filename,
        "file_type": file_extension,
        "file_size_kb": file_size,
        "file_hash_md5": file_hash
    }

def process_directory(directory):
    """
    Recursively processes all files in the given directory,
    returning a list of metadata dictionaries.
    """
    metadata_list = []
    for root, _, files in os.walk(directory):
        for file in files:
            full_path = os.path.join(root, file)
            metadata_list.append(get_file_metadata(full_path))
    return metadata_list

def main():
    if len(sys.argv) != 2:
        print("Usage: python script.py <directory_path>")
        sys.exit(1)
    
    directory = sys.argv[1]
    
    if not os.path.isdir(directory):
        print(f"Error: {directory} is not a valid directory.")
        sys.exit(1)
    
    # process the directory recursively and gather metadata
    metadata = process_directory(directory)
    
    # create DataFrame and save metadata to an Excel file in the provided directory
    df = pd.DataFrame(metadata)
    output_file = os.path.join(directory, "metadata.xlsx")
    df.to_excel(output_file, index=False)
    
    print(f"Metadata saved to {output_file}")

if __name__ == "__main__":
    main()
