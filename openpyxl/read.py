from openpyxl import load_workbook

excel_file = 'FoundHouse.xlsx'    #excel file path
wb = load_workbook(excel_file)

ws = wb.active

for row in ws.iter_rows(values_only=True):
    print(row)  # Print the row values
