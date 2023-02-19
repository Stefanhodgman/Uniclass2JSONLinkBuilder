import os
import pandas as pd
import json
import tkinter as tk
from tkinter import filedialog, messagebox

# Create a Tkinter root window to open a file dialog and select a directory
root = tk.Tk()
root.withdraw()

# Open a file dialog to select a directory
directory = filedialog.askdirectory(title="Select a directory")

# Create a list to store the processed codes dictionaries
all_codes_dict = []

# Iterate over each file in the directory
for filename in os.listdir(directory):
    # Check if the file is an Excel file
    if filename.endswith(".xlsx") or filename.endswith(".xls"):
        filepath = os.path.join(directory, filename)

        # Display a popup asking the user whether to skip the first two lines of the Excel file
        skip_first_two_lines = messagebox.askyesno(title="Skip lines", message="Do you want to skip the first two lines of the Excel file?")

        # Load Excel file into a pandas dataframe
        if skip_first_two_lines:
            df = pd.read_excel(filepath, sheet_name=0, usecols=["Code"], skiprows=2)
        else:
            df = pd.read_excel(filepath, sheet_name=0, usecols=["Code"])

        # Create an empty dictionary to store the codes
        codes_dict = {}

        # Iterate over each row in the dataframe
        for index, row in df.iterrows():
            code = row["Code"]
            parent_code = code[:-2]  # Assume that the parent code is all but the last two characters of the code
            parent_code = parent_code.rstrip("_")  # Remove any trailing underscores from the parent code

            # If the parent code is not in the dictionary, add it as a key with an empty list as the value
            if parent_code not in codes_dict:
                codes_dict[parent_code] = {"parentCode": parent_code, "childCodes": ""}

            # Add the code to the list of child codes for the parent code, separated by "#"
            codes_dict[parent_code]["childCodes"] += f"{code}#"

        # Remove the trailing "#" from each list of child codes
        for code in codes_dict.values():
            code["childCodes"] = code["childCodes"][:-1]

        # Add the processed codes dictionary to the list of all codes dictionaries
        all_codes_dict.append(codes_dict)

        # Save the processed codes dictionary to a JSON file with the same name as the input file
        output_filename = os.path.splitext(filename)[0] + ".json"
        output_filepath = os.path.join(directory, output_filename)
        with open(output_filepath, "w") as json_file:
            json.dump(list(codes_dict.values()), json_file, separators=(',', ':'), indent=2)

# Merge all codes dictionaries into one and save it to a JSON file in the directory
merged_codes_dict = {}
for codes_dict in all_codes_dict:
    merged_codes_dict.update(codes_dict)

merged_codes_list = list(merged_codes_dict.values())
output_filepath = os.path.join(directory, "all_codes.json")
with open(output_filepath, "w") as json_file:
    json.dump(merged_codes_list, json_file, separators=(',', ':'), indent=2)
