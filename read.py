from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import StringFilter

import pandas as pd

# Load the Excel file
book = load_workbook("FoundHouse.xlsx")  # Load the excel file

def search_single_value(sheet_name, column_letter, target):
    
    if sheet_name and column_letter and target:
        sheet = book[sheet_name]
        column = sheet[column_letter] 
        if target.isdigit() and int(target) > 0:
            target = int(target)
        for cell in column:
            if str(cell.value).casefold().strip() == str(target).casefold().strip() or cell.value == target:
                return f'Found {target} in cell {column_letter}{cell.row}'  # Return the cell location if target is found
    return None

def search_in_workbook(sheet_name, targets):
    filter_match_cell = []  # this is an empty list that will store all matching filter
    # Access the sheet
    sheet = book[sheet_name]
    # Search for each target in all columns and print the entire row
    result = ""
    for column in sheet.columns: 
        for cell in column: 
            for target in targets:
                if str(cell.value).casefold().strip() == str(target.casefold()).strip() or cell.value == target:
                    result += f'Found {target} in cell {cell.column_letter}{cell.row}\n'  # Print the target cell location
                    row = sheet[cell.row]  # Get the entire row where the target is found
                    for cell in row: #this will go through each cell in the row
                        result += str(cell.value) + "\t"  # Print all values in the row on the same line
                    result += "\n\n"
    return result

# Kadima's