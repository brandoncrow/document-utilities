# Document Utility Scripts

A collection of lightweight Python scripts for managing, copying, renaming, and moving document files using metadata from Excel.

## Scripts Overview

### 1. `copy_docs_to_subfolder.py`
Copies documents from a source folder into structured subfolders based on metadata (EntityType and Number) from an Excel file. Logs missing and error files.

### 2. `file_copy.py`
Copies and **renames** files using mappings defined in an Excel file. Produces an updated Excel with new paths and an optional error log.

### 3. `move_renamed.py`
Moves already-renamed files from a staging directory to a target location based on an Excel column listing the expected renamed files.

### 4. `get_metadata.py`
Recursively collects file metadata (path, size, hash, type) and outputs it as an Excel sheet. Useful for audit or tracking.

## Setup

```bash
pip install -r requirements.txt
```

Each script is designed to run independently and is executed via the command line (except `copy_docs_to_subfolder.py`, which uses hardcoded paths).

## Usage Examples

```bash
python file_copy.py "mapping.xlsx" "D:/renamed_docs" "OriginalPath" "NewFileName"
python move_renamed.py "updated.xlsx" "D:/staging" "D:/final" "NewFileName"
python get_metadata.py "C:/Documents"
```

## Notes

- All scripts assume files are small enough to handle locally.
- Logs are written for missing or failed copies where applicable.
