import sys
from PyQt5.QtWidgets import QApplication, QLabel, QWidget, QVBoxLayout, QLineEdit, QPushButton, QComboBox, QTableWidget, QTableWidgetItem, QTextEdit
from openpyxl import load_workbook
import pandas as pd
from read import search_single_value, search_in_workbook
import rows

class ourGUI(QWidget):
    def __init__(self):
        super(ourGUI, self).__init__()

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
        self.column_input.setPlaceholderText("Enter the letter name of the column you wish to search in:")
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

        self.results_table = QTableWidget()
        layout.addWidget(self.results_table)

        # Add rows and columns buttons
        self.add_rows_button = QPushButton("Add Rows")
        self.add_rows_button.clicked.connect(self.add_rows_button_clicked)
        layout.addWidget(self.add_rows_button)

        self.add_columns_button = QPushButton("Add Columns")
        self.add_columns_button.clicked.connect(self.add_columns_button_clicked)
        layout.addWidget(self.add_columns_button)

        # Remove rows and columns buttons
        self.row_input = QLineEdit()
        self.row_input.setPlaceholderText("If you want to remove a row, type its number here:")
        layout.addWidget(self.row_input)

        self.remove_rows_button = QPushButton("Remove Rows")
        self.remove_rows_button.clicked.connect(self.remove_rows_button_clicked)
        layout.addWidget(self.remove_rows_button)

        self.col_input = QLineEdit()
        self.col_input.setPlaceholderText("If you want to remove a column, type its letter here:")
        layout.addWidget(self.col_input)

        self.remove_columns_button = QPushButton("Remove Columns")
        self.remove_columns_button.clicked.connect(self.remove_columns_button_clicked)
        layout.addWidget(self.remove_columns_button)

        # Result label
        self.result_label = QTextEdit()
        layout.addWidget(self.result_label)

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
        pass

    def add_columns_button_clicked(self):
        pass

    def remove_rows_button_clicked(self):
        pass

    def remove_columns_button_clicked(self):
        sheet_name = self.sheet_input.currentText()
        letter = self.col_input.text()
        result = rows.remove_column(sheet_name, letter)
        self.result_label.setText(result)

# Kadima's