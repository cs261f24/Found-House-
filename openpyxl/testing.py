import openpyxl

def search_pet_info(file_path, search_name):
    # Load the workbook and select the active sheet
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Iterate through each row in the sheet
    for row in sheet.iter_rows(min_row=2, values_only=True):  # min_row=2 to skip header
        name = row[0]
        pet_name = row[2]
        
        # Check if the search_name matches either the name or pet name
        if search_name.lower() in (name.lower(), pet_name.lower()):
            # Print all values in the row
            print("Found entry:")
            for cell in row:
                print(cell)
            print("\n")  # Add a new line for better readability

# Example usage
file_path = 'FoundHouse.xlsx'  # Update with your file path
search_name = input("Enter a name or pet name to search: ")
search_pet_info(file_path, search_name)
