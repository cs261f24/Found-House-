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
# Function to search and print a single value in a column
def search_single_value():
    sheet_name = input("Enter the sheet you want to search: ")
    column_letter = input("Enter the column you want to search: ")
    target = input("Enter what you want to search for: ")

    # Convert target to integer if it is a positive number
    if target.isdigit() and int(target) > 0: 
        target = int(target)

    # Access the sheet and column
    sheet = book[sheet_name]
    column = sheet[column_letter]

    # Search for the target in the column
    for cell in column:
        if str(cell.value).casefold().strip() == str(target.casefold()).strip() or cell.value == target:
            print(f'Found {target} in cell {column_letter}{cell.row}')  # Print cell location if target is found
        

# Function to search for a target and print the entire row associated with it
def search_in_workbook():
    sheet_name = input("Enter the sheet you want to search: ")
    column_letter = input("Enter the column you want to search: ")
    target = input("Enter what you want to search for: ")

    # Access the sheet and column
    sheet = book[sheet_name]
    column = sheet[column_letter]

    # Search for the target in the column and print the entire row
    for cell in column:
        if str(cell.value).casefold().strip() ==str(target.casefold()).strip() or cell.value == target:
            print(f'Found {target} in cell {column_letter}{cell.row}')  # Print the target cell location
            row = sheet[cell.row]  # Get the entire row where the target is found
            for cell in row:
                print(cell.value, end="\t")  # Print all values in the row on the same line
# Function to allow the user to choose between two search options
def option_search():
    print("Search for a target and print associated values by pressing 1\n")
    print("Search a single target by pressing 2\n")
    option = input("Enter 1 or 2: ")

    if option == "1":
        search_in_workbook()
    elif option == "2":
        search_single_value()
    else:
        print("Invalid input. Please enter either 1 or 2.")

# Call the function to start the search
option_search()
