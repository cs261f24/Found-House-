from openpyxl import load_workbook

# Load the Excel file
excel_file = "FoundHouse.xlsx"   # Path to the Excel file
wb = load_workbook(excel_file)

# Get the active worksheet
ws = wb.active

for row in ws.iter_rows(values_only=True):
    print(row)
