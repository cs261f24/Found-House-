import pandas as pd
import numpy as np
# import openpyxl as op
# from openpyxl import workbook
from openpyxl import load_workbook

# SETS DF TO READ THE SPREADSHEET
df = pd.read_excel(''' input excel spreadsheet file here ''')

# REMOVES NAN FROM BLANK COLUMNS
df = df.fillna('')
df.dropna(inplace=True)
df.to_excel(''' input excel spreadsheet file here ''', index=False)
df1 = df.replace(np.nan, '', regex=True)

# LOADS THE WORKBOOK IF SHEET IS USED
excelBook = load_workbook(''' input spreadsheet file here ''')

writer = pd.ExcelWriter(''' input spreadsheet file here ''', engine='openpyxl')

with pd.ExcelWriter(''' input spreadsheet file here ''') as writer:
    # Saving my the excel spreadsheet as a base.
    writer.book = excelBook
    writer.sheets = dict((ws.title, ws) for ws in excelBook.worksheets)

    # Saves the workbook
    writer.save()


# MAKES A DATAFRAME
print("------------")

# PRINTS WHOLE LIST USING TO_STRING
print(df.to_string())
print("------------")

print(df)
