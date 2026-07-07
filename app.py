import pandas as pd
import pprint
import glob
import os
from pocreater import createPoexcel

output_folder = "PO_Excel_Files"

os.makedirs(output_folder, exist_ok=True)

print('Select action,')
print('Create All Item PO Excel File (Enter: 1 )')
print('Create Excel files for departments (Enter: 2)')
print('-- Press 3 for Exit --')

try:
    action = input()
    action = int(action)
    while(True):

        if action == 1:
            print('-------------------Create PO Excel----------------------')
            createPoexcel(output_folder)
            continue
        elif action == 2:
            print('-------------------Filter PO for departments---------------')
            continue
        
        elif action == 3:
            break

        else:
            print('invalid input')
            break                           
        
except:
    print('error')
# print(type(action))





