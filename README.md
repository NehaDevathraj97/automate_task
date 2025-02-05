# Excel Data Processing Script

## Overview
This script processes Excel files from the specified input folder, extracts data based on a 'Date' column, and consolidates it into a single Excel file. Processed files are then moved to a separate folder.

## Features
- Reads all `.xlsx` files from the input folder
- Checks for a 'Date' column and converts it to datetime format
- Extracts data for the latest month or user-specified month
- Saves the filtered data into a consolidated Excel file
- Moves processed files to a `processed` folder
- Provides interactive user input using Tkinter
- Includes error handling for missing files, invalid date formats, and missing columns

## Requirements
Ensure you have Python installed (preferably Python 3.8 or newer). Required dependencies are listed in `requirements.txt`.

### Install Dependencies
Run the following command to install necessary Python libraries:
```sh
pip install -r requirements.txt
```

## Usage
1. Place your Excel files in the `Input folder`.
2. Run the script using the provided `run_script.bat` file or manually:
   ```sh
   python script.py
   ```
3. The script will prompt for a month selection (leave blank to use the latest month).
4. The processed data will be saved to:
   ```
   consolidated.xlsx
   ```
5. All processed files will be moved to the `processed` folder.

## File Structure
```
project-folder/
│-- automate.py  # Main script
│-- requirements.txt  # Dependencies
│-- run_script.bat  # Batch file to run the script
│-- Input folder/  # Folder containing Excel files
│-- processed/  # Folder where processed files will be moved
```

## Error Handling
- **Missing Input Files:** If no Excel files are found in the input folder, an error message is displayed.
- **Missing 'Date' Column:** If the required 'Date' column is missing from any file, a warning is logged, and the file is skipped.
- **Invalid Dates:** If the 'Date' column contains invalid entries, they are coerced into `NaT` (Not a Time), and only valid dates are used.
- **No Valid Data:** If no valid data remains after filtering, a warning is displayed, and the script exits.
- **File Processing Errors:** If an error occurs while reading a file, the script logs the error and continues processing other files.

## requirements.txt
```
pandas
glob2
openpyxl
tkinter
```

## run_script.bat
Create a `run_script.bat` file with the following content:
```
@echo off
python automate.py
pause
```
Double-click `run_script.bat` to execute the script.

## Notes
- The script only processes files with valid date entries.
- If no valid dates are found, a warning will be displayed.
