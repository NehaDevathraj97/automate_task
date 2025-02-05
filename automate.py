import pandas as pd
import glob
import shutil
import os
import tkinter as tk
from tkinter import simpledialog, messagebox

# Define folders
input_folder = "Input folder"  # Ensure this is correct
processed_folder = "processed"
os.makedirs(processed_folder, exist_ok=True)

# Get all Excel files
all_files = glob.glob(os.path.join(input_folder, "*.xlsx"))

if not all_files:
    messagebox.showerror("Error", "No Excel files found in the input folder!")
    exit()

print("Files found:", all_files)

# Read data
dfs = []
for file in all_files:
    try:
        df = pd.read_excel(file)
        print(f"Processing: {file} | Columns: {df.columns.tolist()}")
        if "Date" in df.columns:
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            valid_dates = df["Date"].dropna()
            print("Valid dates found:", valid_dates.head())
            if not valid_dates.empty:
                dfs.append(df)
        else:
            print(f"Skipping {file}: 'Date' column not found.")
    except Exception as e:
        print(f"Error processing {file}: {e}")

if not dfs:
    messagebox.showerror("Error", "No valid data found in the Excel files!")
    exit()

# Combine data
combined_df = pd.concat(dfs, ignore_index=True)
latest_date = combined_df["Date"].max()
# Get latest year and month
latest_year, latest_month = latest_date.year, latest_date.month
print("Latest available date:", latest_date)

if pd.isna(latest_date):
    messagebox.showerror("Error", "No valid dates found in the dataset!")
    exit()

# Tkinter input for month selection
root = tk.Tk()
root.withdraw()  # Hide main window

#user_year = simpledialog.askinteger("Input", "Enter Year (YYYY) or leave blank for latest:")
#user_month = simpledialog.askinteger("Input", "Enter month (1-12) or leave blank for latest:")

# Prompt user with default values
#user_year = simpledialog.askstring("Input", f"Enter Year (YYYY) or leave blank for {latest_year}:", initialvalue=str(latest_year))
user_month = simpledialog.askstring("Input", f"Enter month (1-12) or leave blank for {latest_month}:", initialvalue=str(latest_month))
if user_month is None:  # User pressed cancel
    messagebox.showinfo("Exit", "Operation cancelled by the user.")
    exit()

# Convert inputs to integers if provided, otherwise use defaults
#user_year = int(user_year) if user_year and user_year.isdigit() else latest_year
user_month = int(user_month) if user_month and user_month.isdigit() else latest_month

print(f"Using Year: {latest_year}, Month: {user_month}")


if user_month is None or user_month not in range(1, 13):
    latest_year, latest_month = latest_date.year, latest_date.month
    print(f"Using latest month: {latest_month}")
else:
    latest_year, latest_month = latest_year, user_month
    print(f"Using user-selected month: {latest_month}")

latest_month_data = combined_df[
    (combined_df["Date"].dt.year == latest_year) & (combined_df["Date"].dt.month == latest_month)
]

if latest_month_data.empty:
    messagebox.showwarning("Warning", f"No data available for month {latest_month}!")
else:
    output_path = "consolidated.xlsx"
    latest_month_data.to_excel(output_path, index=False)
    print(f"Consolidated file created: {output_path}")

    # Move processed files
    for file in all_files:
        destination = os.path.join(processed_folder, os.path.basename(file))
        print(f"Moving {file} to {destination}")
        shutil.move(file, destination)
    print("Files moved to 'processed' folder.")
    messagebox.showinfo("Success", "Processing completed successfully!")
