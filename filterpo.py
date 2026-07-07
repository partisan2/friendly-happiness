import pandas as pd
import pprint
import glob
import os

def filteritems():
    excel_files = glob.glob("./PO_EXCEL_FILES/*.xlsx")

    if not excel_files:
        print("No Excel files found!")
    else:
        print("\n--- Available Excel Files ---")
        # 2. List them out with numbers
        for index, filename in enumerate(excel_files):
            print(f"[{index}] {filename}")

            try:
                choice = int(input("\nEnter the number of the file you want to process: "))
                selected_file = excel_files[choice]
                print(f"You selected: {selected_file}\n")
                
                # Now pass 'selected_file' to your extractor script
                data = pd.read_excel(selected_file, header=None)
                
            except (ValueError, IndexError):
                print("Invalid selection. Please run the script again and choose a valid number.")

    result_dict = {}
    headers = []

    for index, row in data.iterrows():
        print(f"Processing row {index}: {row.values}")

        if "Item Code" in row.values:
            current_location = str(row[0]).strip()
            result_dict[current_location] = []
            headers = list(row.values)
            headers[0] = "Location"
            continue

        if not headers:
            print("Headers not found yet, skipping row.")
            continue

        if row.isna().all():
            continue


