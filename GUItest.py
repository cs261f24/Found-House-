from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QLineEdit, QPushButton, QComboBox
import sys
import openpyxl


class Main(QWidget):
    def __init__(self):
        super(Main, self).__init__()
        self.setWindowTitle("Found House")
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Search input fields
        self.sheet_input = QComboBox()
        layout.addWidget(self.sheet_input)

        self.column_input = QLineEdit() 
        self.column_input.setPlaceholderText("Enter column name")
        layout.addWidget(self.column_input)
        
        self.target_input = QLineEdit()
        self.target_input.setPlaceholderText("Enter what you want to search for: ")
        layout.addWidget(self.target_input)
        
        self.single_search_button = QPushButton("Search Single Value")
        layout.addWidget(self.single_search_button)
        
        self.multi_search_button = QPushButton("Filter by Multiple values")
        layout.addWidget(self.multi_search_button)

        # Table widget to display data
        self.table_widget = QTableWidget()
        layout.addWidget(self.table_widget)

        # Load workbook and set default sheet
        try:
            self.workbook = openpyxl.load_workbook("openpyxl/FoundHouse.xlsx")
        except FileNotFoundError:
            print("Error: Excel file not found. Please ensure the file path is correct.")
            sys.exit()
            
   
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Main()
    window.showMaximized()
    app.exec_()
