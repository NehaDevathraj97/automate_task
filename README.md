# Excel Processing with Tkinter & Pandas

This Python script processes multiple Excel files from an input folder, extracts date-related information, and consolidates data based on user-specified or the latest month. It also moves processed files to a separate folder.

## Features
- Reads all Excel files from the input folder.
- Checks for a `Date` column and ensures valid date conversion.
- Uses Tkinter to prompt the user for a month or defaults to the latest month.
- Combines data from all valid files and filters it based on the selected month.
- Saves the consolidated data to an output Excel file.
- Moves processed files to a `processed` folder.
- Displays user-friendly alerts for errors and completion.

## Requirements
Ensure you have the following Python libraries installed:

```bash
pip install pandas openpyxl tkinter
```

## How to Use
1. Place your `.xlsx` files in the specified `Input folder`.
2. Run the script.
3. When prompted, enter a month (1-12) or leave blank to use the latest month found in the data.
4. When prompted, enter an year (YYYY) or leave blank to use the latest year found in the data.
5. The script will generate a consolidated Excel file and move processed files.
6. Check the output folder for `consolidated.xlsx`.

## Error Handling
- Displays an alert if no files are found.
- Skips files missing the `Date` column.
- Notifies the user if no valid data exists for the selected month.

## File Output
The processed data is saved as:
```
consolidated.xlsx
```

## Author
This script was developed to automate Excel data consolidation using Python and Tkinter.
