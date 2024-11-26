import sys
from PyQt5.QtWidgets import QApplication, QLabel, QWidget, QVBoxLayout, QLineEdit, QPushButton, QComboBox, QTableWidget, QTableWidgetItem, QTextEdit, QInputDialog
from openpyxl import load_workbook
import pandas as pd
from read import search_single_value, search_in_workbook
#from rows import add_column, remove_column, add_row, remove_row
class Main(QWidget):
    def __init__(self):
        super(Main, self).__init__()

        # Set up the layout
        layout = QVBoxLayout()
        self.setWindowTitle("Found House - For the Pets")
        self.setGeometry(100, 100, 900, 600)  # Set the window size
        self.setLayout(layout) 

        # Selecting the respective sheets
        self.sheet_input = QComboBox()
        sheets = load_workbook("FoundHouse.xlsx").sheetnames
        self.sheet_input.addItems(sheets)
        layout.addWidget(self.sheet_input) 

        # Selecting animals
        self.animal_input_label = QLabel("Select Species:")
        self.animal_input = QComboBox()
        self.animal_input.addItems(["Dog", "Cat", "Others"])
        layout.addWidget(self.animal_input_label)
        layout.addWidget(self.animal_input)

        self.column_input = QLineEdit()
        self.column_input.setPlaceholderText("Enter column name")
        layout.addWidget(self.column_input)

        # Target search input field
        self.target_input = QLineEdit()
        self.target_input.setPlaceholderText("Enter what you want to search for:")
        layout.addWidget(self.target_input)

        self.single_search_button = QPushButton("Search Single Value")
        self.single_search_button.clicked.connect(self.search_single_value_button_clicked)
        layout.addWidget(self.single_search_button)

        self.filter_button = QPushButton("Filter by Multiple values")
        self.filter_button.clicked.connect(self.search_in_workbook_button_clicked)
        layout.addWidget(self.filter_button)

         
        self.result_label = QTextEdit()
        layout.addWidget(self.result_label) 

        # Add rows and columns
        self.add_rows_button = QPushButton("Add Rows")
        #adjust size of button
        sheet_name = self.sheet_input.currentText()
        self.add_rows_button.setFixedWidth(150)
        #self.add_rows_button.clicked.connect(self.add_rows_button_clicked)
        layout.addWidget(self.add_rows_button)

        self.add_columns_button = QPushButton("Add Columns")
        self.add_columns_button.setFixedWidth(150)
        layout.addWidget(self.add_columns_button)

        # Remove rows and columns
        self.remove_rows_button = QPushButton("Remove Rows")
        self.remove_rows_button.setFixedWidth(150)
        layout.addWidget(self.remove_rows_button)

        self.remove_columns_button = QPushButton("Remove Columns")
        self.remove_columns_button.setFixedWidth(150)
        layout.addWidget(self.remove_columns_button)

        # Result label
        self.results_table = QTableWidget()
        layout.addWidget(self.results_table) 
      

        try:
            self.workbook = load_workbook("FoundHouse.xlsx")
            self.sheet = self.workbook.active
        except FileNotFoundError:
            print("Error: Excel file not found. Please ensure the file path is correct.")
            sys.exit()

    def search_single_value_button_clicked(self):
        sheet_name = self.sheet_input.currentText()
        column_letter = self.column_input.text()
        target = self.target_input.text()
        result = search_single_value(sheet_name, column_letter, target)
        if result is not None:
            self.result_label.setText(result)
        else:
            self.result_label.setText("Target not found")
            
    def search_in_workbook_button_clicked(self):
        sheet_name = self.sheet_input.currentText()
        targets = self.target_input.text().split(',')
        result = search_in_workbook(sheet_name, targets)
        if result:
            self.result_label.setText(result)
        else:
            self.result_label.setText("not found")
            
    def add_rows_button_clicked(self):
        sheet_name = self.sheet_input.currentText()
        user_input = QInputDialog.getInt(self, "Add Rows", "Enter the number of rows to add:")
        
        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Main()
    window.showMaximized()  # Maximizes the window
    app.exec_()