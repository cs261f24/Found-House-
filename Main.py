import sys
from openpyxl import load_workbook, Workbook
import pandas
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import *
import GUIexperimentation
import read

def main():
    app = QApplication(sys.argv)
    window = GUIexperimentation.ourGUI()
    window.showMaximized()  # Maximizes the window
    app.exec_()

if __name__ == '__main__':
    main()

# Frazee's