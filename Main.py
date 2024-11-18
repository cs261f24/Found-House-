import sys
from openpyxl import load_workbook, Workbook
import pandas
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import *
import GUIexperimentation
from read import search_in_workbook, option_search

def main():
    book = load_workbook("FoundHouse.xlsx")
    result = search_in_workbook(book)
    print(result)

if __name__ == '__main__':
    main()
