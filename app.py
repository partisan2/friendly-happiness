import pandas as pd
import pprint
import glob
import os
from pocreater import createPoexcel
from filterpo import filteritems

output_folder = "PO_Excel_Files"

os.makedirs(output_folder, exist_ok=True)


while True:
    print("\nSelect action")
    print("1. Create All Item PO Excel File")
    print("2. Create Excel files for departments")
    print("3. Exit")

    try:
        action = int(input("> "))

        if action == 1:
            print("-------------------Create PO Excel----------------------")
            createPoexcel(output_folder)

        elif action == 2:
            print("-------------------Filter PO for departments---------------")
            filteritems()
            print("Finished!")

        elif action == 3:
            break

        else:
            print("Invalid input")

    except Exception as e:
        print("Error:", e)
# print(type(action))





