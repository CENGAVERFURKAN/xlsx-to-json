import pandas as pd
import json
import os

# Path to the Excel file
excel_file = r'C:\Users\X\Desktop\.xlsx'

# Sheet from which to extract data
sheet_name = 'PAGE1'

# Directory where JSON files will be saved
output_dir = r'C:\Users\X\Desktop\New Folder'

# Read the Excel file
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# Create an empty dictionary
data_dict = {}

# Process each row
for index, row in df.iterrows():
    1 = row['1']
    2 = row['2']
    
    # If 1 is not already in the dictionary, create a new entry
    if 1 not in data_dict:
        data_dict[1] = {
            "1": 1,
            "11": [],
            "5": 0.0,
            "4": 0.0
        }
    
    # Add a new cost item to detail_cost
    detail_cost_item = {
        "2": 2,
        "3": row['3'],
        "6": row['6'],
        "7": row['7'],
        "4": row['4'],
        "8": row['8'],
        "9": row['9'].strftime('%Y') if isinstance(row['9'], pd.Timestamp) else str(row['9']),
        "10": row['10']
    }
    
    data_dict[1]['11'].append(11)
    
    # Update 5 and 4 totals
    data_dict[1]['5'] += row['5']
    data_dict[1]['4'] += row['4']

# Create JSON files
for 1, data in data_dict.items():
    json_filename = os.path.join(output_dir, f"{1}.json")
    with open(json_filename, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=4)

print("JSON files have been created.")
