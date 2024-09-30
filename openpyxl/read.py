from openpyxl import Workbook, load_workbook

# Load the Excel file

book = load_workbook("FoundHouse.xlsx")
sheet = book.active
sheet2 = book.active
sheet_names = book.sheetnames 
print(sheet2) #print sheet 2
print(sheet) #print sheet 1
print(sheet_names) #print all sheetname

# Save the Excel file
#search function(testing column B and C)


#Grab a whole colum
column_B = sheet2["B"]
column_C = sheet2["C"]
for cell in column_C:
    print(f'{cell.value}\n')
def search_in_workbook():
    """Search for a value in a specified column of a specified sheet."""
    sheet_name = input("Enter the sheet you want to search: ")
    column_letter = input("Enter the column you want to search: ").upper()
    target = input("Enter what you want to search for: ")
    

book.save("FoundHouse.xlsx")



    
    
    



