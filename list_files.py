import os
from openpyxl import Workbook

# specify directory path
directory_path = input("Enter the directory path: ")

# create excel workbook and select worksheet
wb = Workbook()
ws = wb.active
ws.title = "File Names"

# add header
ws.append(["File Name"])

# loop through files in dir
for filename in os.listdir(directory_path):
    filename_path = os.path.join(directory_path, filename)
     # Include files and directories
    if os.path.isfile(filename_path) or os.path.isdir(filename_path):
        ws.append([filename])
        
# save excel file
output_path = input("Enter the output path (e.g. text.xlsx) ")
wb.save(output_path)

print(f"File names have been saved to {output_path}")