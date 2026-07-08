import pandas as pd
import pprint
import glob
import os

def createPoexcel(output_folder):
    excel_files = glob.glob("*.xlsx")

    if not excel_files:
        print("No Excel files found!")
    else:
        print("\n--- Available Excel Files ---")
        # 2. List them out with numbers
        for index, filename in enumerate(excel_files):
            print(f"[{index}] {filename}")
        
        # 3. Prompt the user for a selection
        try:
            choice = int(input("\nEnter the number of the file you want to process: "))
            selected_file = excel_files[choice]
            print(f"You selected: {selected_file}\n")
            
            # Now pass 'selected_file' to your extractor script
            data = pd.read_excel(selected_file, header=None)
            
        except (ValueError, IndexError):
            print("Invalid selection. Please run the script again and choose a valid number.")


    result_dict = {}
    current_location = None
    headers = []

    for index, row in data.iterrows():
        #print(f"Processing row {index}: {row.values}")
        
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

        if current_location:
            row_data = {}
            for col_idx, header_name in enumerate(headers):
                # Skip the first column (where 'nan' sits in data rows)
                if col_idx == 0 or pd.isna(header_name):
                    continue

                row_data[header_name] = row[col_idx]

            # Only append if the row actually has valid data values
            if any(pd.notna(val) for val in row_data.values()):
                result_dict[current_location].append(row_data)
    # import pprint

    # pprint.pprint(result_dict)

    locations = sorted(list(result_dict.keys()))

    # print(locations)

    pivoted_items = {}

    for location_name, items in result_dict.items():
        for item in items:
            code = item.get("Item Code")
            desc = item.get("Item Description")
            qty = item.get("Qty")
            
            # Skip invalid or empty items
            if pd.isna(code):
                continue
                
            # Standardize data types (convert code to string or int cleanly)
            code = int(code) if isinstance(code, float) and code.is_integer() else code

            # If we haven't seen this item code yet, initialize its entry
            if code not in pivoted_items:
                pivoted_items[code] = {
                    "Item Code": code,
                    "Item Description": desc
                }
                # Initialize all known locations with 0 quantity
                for loc in locations:
                    pivoted_items[code][loc] = 0

            # Add the quantity to the specific location slot (handles duplicates safely if any)
            if pd.notna(qty):
                pivoted_items[code][location_name] += int(qty)

    # 3. Flatten the dictionary into a list and calculate T.QTY (Grand Total)
    final_list = []
    for code, details in pivoted_items.items():
        # Sum up the values of all location columns for this item
        total_qty = sum(details[loc] for loc in locations)
        details["T.QTY"] = total_qty
        final_list.append(details)

    # 4. View your new structured dictionary list
    # pprint.pprint(final_list)

    df_final = pd.DataFrame(final_list)

    # Reorder columns explicitly to ensure T.QTY is at the very end
    column_order = ["Item Code", "Item Description"] + locations + ["T.QTY"]
    df_final = df_final[column_order]

    # Save to a new Excel file
    output_filename = input("Enter the name for the output Excel file (without extension): ")
    full_output_path = os.path.join(output_folder, output_filename)
    return df_final.to_excel(f"{full_output_path}.xlsx", index=False)