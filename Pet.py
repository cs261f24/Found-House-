import pandas as pd
from openpyxl import load_workbook

# SETS DF TO READ THE SPREADSHEET
df = pd.read_excel('''Insert Spreadsheet Here''')

# REMOVES NAN FROM BLANK COLUMNS
df = df.fillna('')

# LOADS THE WORKBOOK IF SHEET IS USED WITHOUT DELETING SHEETS
excelBook = load_workbook('''Insert Spreadsheet Here''')

# USES EXCELWRITER IN APPEND MODE WITH SHEET REPLACEMENT IF IT EXISTS
with pd.ExcelWriter('''Insert Spreadsheet Here''', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, index=False)

# MAKES A DATAFRAME
print("------------")

# PRINTS WHOLE LIST USING TO_STRING

print(df.to_string())

print("------------")
