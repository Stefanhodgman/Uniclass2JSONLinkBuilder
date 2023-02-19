import pandas as pd
import json

# Load Excel file into a pandas dataframe
df = pd.read_excel(r"C:\Uniclass\Uniclass2015_EF.xlsx", sheet_name=0, usecols=["Code"])

# Create an empty dictionary to store the codes
codes_dict = {}

# Iterate over each row in the dataframe
for index, row in df.iterrows():
    code = row["Code"]
    parent_code = code[:-2]  # Assume that the parent code is all but the last two characters of the code
    
    # If the parent code is not in the dictionary, add it as a key with an empty list as the value
    if parent_code not in codes_dict:
        codes_dict[parent_code] = {"parentCode": parent_code, "childCodes": []}
    
    # Add the code to the list of child codes for the parent code
    codes_dict[parent_code]["childCodes"].append(code)

# Convert the dictionary to JSON and save it to a file
with open(r"C:\Uniclass\codes.json", "w") as json_file:
    json.dump(list(codes_dict.values()), json_file, separators=(',', ':'))
