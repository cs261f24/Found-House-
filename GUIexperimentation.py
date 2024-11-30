import sys
from PyQt5.QtWidgets import QApplication, QLabel, QWidget, QVBoxLayout, QLineEdit, QPushButton, QComboBox, QTableWidget, QTextEdit, QSplitter, QFrame
from openpyxl import load_workbook
from read import search_in_workbook, search_single_value
class Main(QWidget):
    def __init__(self):
        super(Main, self).__init__()

        # Set up the layout
        self.setWindowTitle("Found House - For the Pets")
        self.setGeometry(100, 100, 900, 600)  # Set the window size

        # Main splitter for vertical division
        main_splitter = QSplitter()
        main_splitter.setOrientation(1)  # Vertical orientation

        # Left Panel
        left_panel = QFrame()
        left_layout = QVBoxLayout()
        left_panel.setLayout(left_layout)

        # Selecting the respective sheets
        self.sheet_input = QComboBox()
        sheets = load_workbook("FoundHouse.xlsx").sheetnames
        self.sheet_input.addItems(sheets)
        left_layout.addWidget(self.sheet_input)

        # Selecting animals
        self.animal_input_label = QLabel("Select Species:")
        self.animal_input = QComboBox()
        self.animal_input.addItems(["Dog", "Cat", "Others"])
        left_layout.addWidget(self.animal_input_label)
        left_layout.addWidget(self.animal_input)

        self.column_input = QLineEdit()
        self.column_input.setPlaceholderText("Enter column name")
        left_layout.addWidget(self.column_input)

        # Target search input field
        self.target_input = QLineEdit()
        self.target_input.setPlaceholderText("Enter what you want to search for:")
        left_layout.addWidget(self.target_input)

        self.single_search_button = QPushButton("Search Single Value")
        self.single_search_button.clicked.connect(self.search_single_value_button_clicked)
        left_layout.addWidget(self.single_search_button)

        self.filter_button = QPushButton("Filter by Multiple values")
        self.filter_button.clicked.connect(self.search_in_workbook_button_clicked)
        left_layout.addWidget(self.filter_button)

        # Add rows and columns
        self.add_rows_button = QPushButton("Add Rows")
        self.add_rows_button.setFixedWidth(150)
        left_layout.addWidget(self.add_rows_button)

        self.add_columns_button = QPushButton("Add Columns")
        self.add_columns_button.setFixedWidth(150)
        left_layout.addWidget(self.add_columns_button)

        # Remove rows and columns
        self.remove_rows_button = QPushButton("Remove Rows")
        self.remove_rows_button.setFixedWidth(150)
        left_layout.addWidget(self.remove_rows_button)

        self.remove_columns_button = QPushButton("Remove Columns")
        self.remove_columns_button.setFixedWidth(150)
        left_layout.addWidget(self.remove_columns_button)

        # Right Panel
        right_panel = QFrame()
        right_layout = QVBoxLayout()
        right_panel.setLayout(right_layout)

        self.results_table = QTableWidget()
        right_layout.addWidget(self.results_table)

        self.result_label = QTextEdit()
        right_layout.addWidget(self.result_label)

        # Add panels to the splitter
        main_splitter.addWidget(left_panel)
        main_splitter.addWidget(right_panel)

        # Set the splitter as the main layout
        layout = QVBoxLayout(self)
        layout.addWidget(main_splitter)
        self.setLayout(layout)

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
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Main()
    window.show()
    sys.exit(app.exec_())


