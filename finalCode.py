#e2efda


import openpyxl
import re

# Load the Excel file
workbook = openpyxl.load_workbook('/Users/adomoshagan/Desktop/Summer 23/GitHub Portfolio/Scheduling/JYCScheduling/data.xlsx')

# Iterate over each sheet in the workbook
for sheet_name in workbook.sheetnames:
    sheet = workbook[sheet_name]
    
    # Iterate over each cell in the sheet
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and re.match(r'\bavailable\b', str(cell.value), re.IGNORECASE):
                cell.value = sheet.title

# Save the modified workbook
workbook.save('modified_file.xlsx')
