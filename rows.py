import pandas as pd
from openpyxl import Workbook, load_workbook


output_file = "FoundHouse.xlsx"

def load_data():
    sheet = pd.read_excel(output_file)
    return pd.DataFrame(sheet)
sheetdf = load_data()

sheetdf.columns = sheetdf.columns.map(str)

def add_column(sheetdf):
    #name of the column
    column_add = input("What new column would you like to add (if you dont want to add anything, type No) : ")
    #doesnt change anything 
    if column_add.lower() == "no":
        print("You decided not to add a column")
        return sheetdf
    else:
        sheetdf[column_add] = 0 
        #keeps asking the user to input values for each row
        for i in range (len(sheetdf)):
            row_name = sheetdf.iloc[i,0]
            value = input (f"Enter a value for '{row_name}' in the new column '{column_add}': ")
            sheetdf.at[i, column_add] = value
        #saves it to the excel file 
        save_to_excel(sheetdf)
        print ("The new column " + column_add + " was saved to the dataframe")
    return sheetdf

def remove_column(sheetdf):
    #prints the columns available to remove 
    print ("The columns in the dataframe are: ")
    print (sheetdf.columns.tolist())
    #asks the user which column to remove
    column_remove = (input("What column would you like to take out (if you dont want to add anything, type no : "))
    column_remove = column_remove.strip()
    if column_remove.lower() == "no":
        print ("you decided not to remove a column")
        return sheetdf
    #checks if the column exists in the dataframe 
    elif column_remove in sheetdf.columns:
        del sheetdf[column_remove]
        print ("The column " + column_remove + " was found and removed")
    #saves it to the excel file
        save_to_excel(sheetdf)
    else:
        print ("The column was not found")
    return sheetdf


def add_row(sheetdf):
    row_name = input("what new row would you like to add (if you dont want to, type no) : ")
    if row_name.lower() == "no":
        return sheetdf
    else: 
        print ("The columns are " + str(sheetdf.columns.tolist()))
        new_row = [row_name]
        #keeps asking the suer for values to input to each cell
        for column in sheetdf.columns[1:]:
            value = input(f"Enter a value for '{column}' in the new row '{row_name}': ")
            new_row.append(value)
        sheetdf.loc[len(sheetdf)] = new_row  
        #save to excel file  
        save_to_excel(sheetdf)
    print ("The new row " + row_name + " was saved to the dataframe")
    return sheetdf

def remove_row(sheetdf):
    print("The rows available to remove are: ")
    print (sheetdf.iloc[:, 0].astype(str).tolist())
    row_remove = (input("What row would you like to remove(if you dont want to remove a row, type No): "))
    if row_remove.lower() == "no":
        print("You decided not to remove a row")
        return sheetdf
    #removes a row by converting index to string 
    if row_remove in sheetdf.iloc[:, 0].astype(str).values:
        index_to_remove = sheetdf[sheetdf.iloc[:, 0].astype(str) == row_remove].index[0]
        sheetdf.drop(index_to_remove, axis = 0, inplace = True)
        print (f"The row '{row_remove}' was removed")
        #saves to excel file 
        save_to_excel(sheetdf)
        print ("the row was not found in the dataframe")
    return sheetdf

def save_to_excel(sheetdf):
    with pd.ExcelWriter(output_file, engine="openpyxl", mode='w') as writer:  # 'w' mode to overwrite the entire file
        sheetdf.to_excel(writer, index=False)



# keeps asking the user if they would like to edit the file, until the user types "exit"
while True:  
    action = input ("Would you like to add or remove a column or row? Type 'add/remove column' or 'add/remove row' or 'exit' to quit: ")
    if action == "add column":
        sheetdf = add_column(sheetdf)
    elif action == "remove column":
        sheetdf = remove_column(sheetdf)
    elif action == "add row":
        sheetdf = add_row(sheetdf)
    elif action == "remove row":
        sheetdf = remove_row(sheetdf)
    elif action == "exit":
        break 
    else: 
        print ("action not found")

print("The finals dataframe is : ")
print (sheetdf)