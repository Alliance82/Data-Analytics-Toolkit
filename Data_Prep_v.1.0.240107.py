# Created By Alliance82
# Created On 1/7/2024
# Project to help prepare data and clean it then reoutput it
# Currently dynamically reads .xlsx or .csv files and reads them into a DataFrame
import pandas as pd
import numpy as np 
import sys

# Update the file path, file name, extension, and sheet name below 
# The program will determine what type of pandas read commands to execute
file_path = r'Your File Path Here'
file_name = 'Your file name here'
extension = '.xlsx'
sheet = 'Sheet1'

# Produces the full file path including name and extension that will be read into the DataFrame
full_path = file_path + file_name + extension

# Creating a dictionairy of extensions 
file_params = {
    ".xlsx": {"read_func": "read_excel", "read_path": [full_path], "read_sheets": {"sheet_name": sheet}},
    ".xls": {"read_func": "read_excel", "read_path": [full_path], "read_sheets": {"sheet_name": sheet}},
    ".xlsb": {"read_func": "read_excel", "read_path": [full_path], "read_sheets": {"sheet_name": sheet}},
    ".csv": {"read_func": "read_csv", "read_path": [full_path]},
}

# Reading the extension type 
file_info = file_params.get(extension, None)

if file_info:
    # Will try to read the file using pandas with the sheet and if it can't it will drop it
    # This enables looking at files that do not have a sheet like a csv
    try:
        # Dynamically call the appropriate pandas read function
        df = getattr(pd, file_info["read_func"])(*file_info["read_path"], **file_info["read_sheets"])
    except:
        df = getattr(pd, file_info["read_func"])(*file_info["read_path"])
else:
    print(f"Unsupported file extension: {extension}")
    
print(df)    
print(df.columns)

# Removes the whitespace from the Column names
df.columns = df.columns.str.strip()
print(df.columns)

# Outputs an excel file of the same name of the original file, with the cleaned columns
df.to_excel(full_path, index=False, sheet_name=sheet)
