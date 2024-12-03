from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import StringFilter
import pandas as pd
import re

# Load the Excel file
book = load_workbook("FoundHouse.xlsx")  # Load the excel file
def search_single_value(sheet_name, column_header, target):
    result = ""
    if sheet_name and column_header and target:
        sheet = book[sheet_name]
        headers = []
        for cell in sheet[2] or sheet[0]:  # Read the first row to get column headers
            headers.append(cell.value)
        if column_header.isalpha() and len(column_header) == 1: # Check if the column header is a single letter 
            column_letter = column_header.upper()
        else:
            # Map column header to letter
            column_letter = chr(ord('A') + headers.index(column_header))
        
        column = sheet[column_letter] 
        if target.isdigit() and int(target) > 0:
            target = int(target)
        for cell in column:
            if str(cell.value).casefold().strip() == str(target).casefold().strip() or cell.value == target:
                return f'Found {target} in cell {column_letter}{cell.row}'
            row = sheet[cell.row] #get the row
            for cell in row: #loop through the row
                result += str(cell.value) + "\t" #print the value of the cell
            result += "\n"
            # Return the cell location if target is found
    return result
def check_conditions(cell_value, operator, value):
    if operator == '=':
        return str(cell_value).casefold().strip() == str(value).casefold().strip()
    elif operator == '<=':
        return cell_value <= int(value)
    elif operator == '>=':
        return cell_value >= int(value)
    elif operator == '<':
        return cell_value < int(value)
    elif operator == '>':
        return cell_value > int(value)
    else:
        return False

def search_in_workbook(sheet_name, targets):
    conditions = [] #to store the conditions
    for target in targets: 
        match = re.match(r'(\w+)\s*([<>=]+)\s*(\w+)', target) 
        if match is not None:
            column, operator, value = match.groups()
            conditions.append((column, operator, value))
        else:
            # If the target is not a condition, treat it as a single value
            conditions.append((None, None, target))

    sheet = book[sheet_name] #get the sheet name
    headers = [cell.value for cell in sheet[2]]  # Read the first row to get column headers

    result = ""
    for column_index, column in enumerate(sheet.columns): #this just loops through the columns
        column_header = headers[column_index] # grabbing the column header
        for cell in column[2:]:  # Skip the header row
            row_index = cell.row 
            row_values = [cell.value for cell in sheet[row_index]] #getting the values in the row
            match_any_conditions = False
            found_target = None
            for column_name, operator, value in conditions:
                if column_name == column_header:
                    cell_value = cell.value
                    if check_conditions(cell_value, operator, value):
                        match_any_conditions = True
                        found_target = value
                        break #get out of the loop, no need to check the rest when a match is found
                else:
                    # If the condition is a single value, check if it's present in the row
                    if value in row_values: #make sure that the value is in the row
                        match_any_conditions = True 
                        found_target = value 
                        break
            if match_any_conditions == True:
                result += f"Found {found_target} in cell {cell.column_letter}{cell.row}\n"
                row = sheet[cell.row] #get the row
                for cell in row: #loop through the row
                    result += str(cell.value) + "\t" #print the value of the cell
                result += "\n"
                    
    return result
