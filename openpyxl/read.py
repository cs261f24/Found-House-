from openpyxl import Workbook, load_workbook

# Load the Excel file

book = load_workbook("FoundHouse.xlsx")
sheet = book.active

sheet_names = book.sheetnames
print(sheet)
print(sheet_names)
print(sheet["A2"].value)