from openpyxl import Workbook, load_workbook
import pandas as pd
# Load the Excel file

book = load_workbook("FoundHouse.xlsx") #Load the excel file
print(book.sheetnames) #Print all sheetnames
"""

sheet1 = book["Yearly Stats"] #rename sheet1 to Yearly Stats
sheet2 = book["Fake Data Served"] #rename sheet2 to Fake Data Served
sheet3 = book["Fake Data Needs Help"] #rename sheet3 to Fake Data Needs Help
sheet_names = book.sheetnames 
print("This is", sheet1.title) #Print the sheet name
print("This is ", sheet2.title)
print("This is ", sheet3.title)
print(sheet_names) #print all sheetname

# Save the Excel file
#search function(testing column B and C)


#Grab a whole column
column_B = sheet2["B"] #Grab a whole column
column_C = sheet2["C"]

sheet2["B5"].value = 'test'
print(sheet2["B5"].value)

#Grab a single cell and print the value inside of it
for cell in column_C:
    print(f'{cell.value}\n')
"""

def option_search():
    
    pass
"""
def search_integers():
    sheet_name = input("Enter the sheet you want to search: ")
    column_letter = input("Enter the column you want to search: ")
    target = input("Enter what you want to search for: ")
    if target.isdigit() and int(target) > 0: 
        target = int(target)
        sheet = book[sheet_name] 
        column = sheet[column_letter]
        
        for cell in column: 
            if cell.value == target:
                print(f'Found {target} in cell {column_letter}{cell.row}') # Print the cell location for integer values
search_integers()
"""
    
def search_in_workbook():
    sheet_name = input("Enter the sheet you want to search: ")
    column_letter = input("Enter the column you want to search: ")
    target = input("Enter what you want to search for: ")
    if target.isdigit() and int(target) > 0: 
        target = int(target)
        sheet = book[sheet_name] 
        column = sheet[column_letter]
        
        for cell in column: 
            if cell.value == target:
                print(f'Found {target} in cell {column_letter}{cell.row}') # Print the cell location for integer values
    else:
        if sheet_name and column_letter and target:
            sheet = book[sheet_name] # Get the sheet from the workbook
            column = sheet[column_letter] # Get the column from the sheet
            for cell in column: # For each cell in the column and print the value
                if str(cell.value).casefold().strip() in str(target.casefold()).strip() or cell.value == target:
                    print(f'Found {target} in cell {column_letter}{cell.row}') # Print the cell where the target was found
                    row  = sheet[cell.row] # Get the row that the target was found in
                    for cell in row: # For each cell in the row
                        print(cell.value, end="\t") # Print the value and end kis used to make sure the values are printed in the same line 
search_in_workbook() 

