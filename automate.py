import pandas as pd
import glob
import shutil
import os

input_folder = "Input folder"  # Ensure this is correct
processed_folder = "processed"
 
os.makedirs(processed_folder, exist_ok=True)
 
all_files = glob.glob(os.path.join(input_folder, "*.xlsx"))
print("Files found:", all_files)
print(os.getcwd())
dfs = []
for file in all_files:
    try:
        df = pd.read_excel(file)
        print(f"Processing: {file} | Columns: {df.columns.tolist()}")
        if "Date" in df.columns:
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            print("Valid dates found:", df["Date"].dropna().head())
            dfs.append(df)
        else:
            print(f"Skipping {file}: 'Date' column not found.")
    except Exception as e:
        print(f"Error processing {file}: {e}")
 
if dfs:
    combined_df = pd.concat(dfs, ignore_index=True)
    # Get latest date
    latest_date = combined_df["Date"].max()
    print("Latest available date:", latest_date)
 
    latest_year, latest_month = latest_date.year, latest_date.month
    latest_month_data = combined_df[
        (combined_df["Date"].dt.year == latest_year) & (combined_df["Date"].dt.month == latest_month)
    ]
    print("Filtered rows:", len(latest_month_data))
    latest_month_data.to_excel("consolidated.xlsx", index=False)
    print("Consolidated file created: 'consolidated.xlsx'")
 
    for file in all_files:
        destination = os.path.join(processed_folder, os.path.basename(file))
        print(f"Moving {file} to {destination}")
        shutil.move(file, destination)
    print("Files moved to 'processed' folder.")
else:
    print("No valid data to process.")
