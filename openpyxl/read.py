from openpyxl import Workbook, load_workbook

# Load the Excel file

book = load_workbook("FoundHouse.xlsx")
sheet1 = book["Yearly Stats"]
sheet2 = book["Fake Data Served"]
sheet3 = book["Fake Data Needs Help "]
sheet_names = book.sheetnames 
print("This is", sheet1.title)
print("This is ", sheet2.title)
print("This is ", sheet3.title)
print(sheet_names) #print all sheetname

# Save the Excel file
#search function(testing column B and C)


#Grab a whole column
column_B = sheet2["B"]
column_C = sheet2["C"]

sheet2["B5"].value = 'test'
print(sheet2["B5"].value)


for cell in column_C:
    print(f'{cell.value}\n')

def search_in_workbook():
    sheet_name = input("Enter the sheet you want to search: ")
    column_letter = input("Enter the column you want to search: ").upper()
    target = input("Enter what you want to search for: ")
    if sheet_name and column_letter and target:
        sheet = book[sheet_name]
        column = sheet[column_letter]
        for cell in column:
            if cell.value == target:
                print(f'Found {target} in cell {column_letter}{cell.row}')
                
search_in_workbook()
book.save("FoundHouse.xlsx")



    
    
    



