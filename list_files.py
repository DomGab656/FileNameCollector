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
    # Check if it is a file
    if os.path.isfile(os.path.join(directory_path, filename)):
        ws.append([filename])
        
# save excel file
output_path = input("Enter the output path (e.g. text.xlsx) ")
wb.save(output_path)

print(f"File names have been saved to {output_path}")