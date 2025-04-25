"""
move_renamed.py

Moves files from a source directory to a target directory based on renamed file names listed in an Excel sheet.

Usage:
    python move_renamed.py <excel_file> <source_dir> <target_dir> <renamed_file_column>

Example:
    python move_renamed.py "docs_updated.xlsx" "C:/Staging" "D:/Final" "NewFileName"
"""

import os
import sys
import shutil
import pandas as pd


def main():
    if len(sys.argv) != 5:
        print("Usage: python move_renamed.py <excel_file> <source_dir> <target_dir> <renamed_file_column>")
        sys.exit(1)

    excel_path = sys.argv[1]
    source_dir = sys.argv[2]
    target_dir = sys.argv[3]
    column_name = sys.argv[4]

    # Check paths
    if not os.path.exists(excel_path):
        print(f"Error: Excel file not found: {excel_path}")
        sys.exit(1)

    if not os.path.isdir(source_dir):
        print(f"Error: Source directory not found: {source_dir}")
        sys.exit(1)

    os.makedirs(target_dir, exist_ok=True)

    # load Excel file
    df = pd.read_excel(excel_path)

    if column_name not in df.columns:
        print(f"Error: Column '{column_name}' not found in Excel file.")
        print("Available columns:", list(df.columns))
        sys.exit(1)

    # collect renamed file names
    expected_files = set(str(name) for name in df[column_name].dropna())

    moved = 0
    skipped = 0
    for fname in os.listdir(source_dir):
        if fname in expected_files:
            src = os.path.join(source_dir, fname)
            dst = os.path.join(target_dir, fname)
            try:
                shutil.move(src, dst)
                moved += 1
            except Exception as e:
                print(f"Error moving {fname}: {e}")
                skipped += 1

    print(f"Moved {moved} files to {target_dir}.")
    if skipped:
        print(f"Skipped {skipped} files due to errors.")

if __name__ == "__main__":
    main()
