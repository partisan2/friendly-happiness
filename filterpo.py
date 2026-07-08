import glob
import pandas as pd
import json
import os
import pprint
from datetime import datetime

def getPoItems():

    excel_files = glob.glob("./PO_EXCEL_FILES/*.xlsx")

    if not excel_files:
        print("No Excel files found!")
        return

    print("\n--- Available Excel Files ---")
    for index, filename in enumerate(excel_files):
        print(f"[{index}] {filename}")

    try:
        choice = int(input("\nEnter the number of the file you want to process: "))
        selected_file = excel_files[choice]
    except (ValueError, IndexError):
        print("Invalid selection.")
        return

    df = pd.read_excel(selected_file)

    # Find column names dynamically
    code_col = "Item Code"
    desc_col = "Item Description"
    qty_col = "T.QTY"

    result_dict = {}

    for _, row in df.iterrows():
        code = row[code_col]

        # Skip empty rows
        if pd.isna(code):
            continue

        result_dict[str(code)] = {
            "description": row[desc_col],
            "total_qty": int(row[qty_col]) if not pd.isna(row[qty_col]) else 0
        }

    return result_dict


def filteritems():
    data = getPoItems()

    print('start filtering')

    if data is None:
        return None
    
    print("Loaded PO")

    with open("item_list.json", "r") as file:
        item_list = json.load(file)

    # Create lookup: product_id -> category
    product_lookup = {}

    for category, products in item_list["item_list"].items():
        for product in products:
            product_lookup[str(product["product_id"])] = {
                "category": category,
                "product_name": product["product_name"]
            }

    # Group items by category
    category_items = {}

    print("Writing files")

    for code, item in data.items():
        if code not in product_lookup:
            continue

        category = product_lookup[code]["category"]

        if category not in category_items:
            category_items[category] = []

        category_items[category].append({
            "Code": code,
            "Description": item["description"],
            "Product Name": product_lookup[code]["product_name"],
            "Quantity": item["total_qty"]
        })

    # Create output folder
    os.makedirs("output", exist_ok=True)

    now = datetime.now()
    string_date = now.strftime("%Y-%m-%d")

    # Create one Excel file per category
    for category, items in category_items.items():
        df = pd.DataFrame(items)
        filename = f"output/{category}-{string_date}.xlsx"
        df.to_excel(filename, index=False)

        print(f"Created: {filename}")

    print("Done")

    return None

