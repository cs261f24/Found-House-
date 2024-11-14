import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QTextEdit
from openpyxl import load_workbook

# Load the Excel file
book = load_workbook("FoundHouse.xlsx")  # Adjust the path to the correct Excel file

class ExcelSearchApp(QWidget):
    def __init__(self):
        super().__init__()

        # Set up the window layout
        self.setWindowTitle("Excel Search Tool")
        self.setGeometry(100, 100, 500, 300)
        layout = QVBoxLayout()

        # Input fields
        self.sheet_label = QLabel("Enter Sheet Name:")
        self.sheet_input = QLineEdit()

        self.column_label = QLabel("Enter Column Letter:")
        self.column_input = QLineEdit()

        self.target_label = QLabel("Enter Target to Search:")
        self.target_input = QLineEdit()

        # Search buttons
        self.single_search_button = QPushButton("Search Single Value")
        self.single_search_button.clicked.connect(self.search_single_value)

        self.row_search_button = QPushButton("Search and Print Entire Row")
        self.row_search_button.clicked.connect(self.search_in_workbook)

        self.table_list_button = QPushButton("List All Tables in Sheet")
        self.table_list_button.clicked.connect(self.list_tables)

        # Output display
        self.result_display = QTextEdit()
        self.result_display.setReadOnly(True)

        # Add widgets to the layout
        layout.addWidget(self.sheet_label) #
        layout.addWidget(self.sheet_input)
        layout.addWidget(self.column_label)
        layout.addWidget(self.column_input)
        layout.addWidget(self.target_label)
        layout.addWidget(self.target_input)
        layout.addWidget(self.single_search_button)
        layout.addWidget(self.row_search_button)
        layout.addWidget(self.table_list_button)
        layout.addWidget(self.result_display)

        self.setLayout(layout)

    # Function to search and print a single value in a column
    def search_single_value(self):
        sheet_name = self.sheet_input.text()
        column_letter = self.column_input.text()
        target = self.target_input.text()

        try:
            sheet = book[sheet_name]
            column = sheet[column_letter]

            for cell in column:
                if str(cell.value).casefold().strip() == str(target.casefold()).strip() or cell.value == target:
                    result = f'Found {target} in cell {column_letter}{cell.row}\n'
                    self.result_display.append(result)
                    return
            self.result_display.append(f"{target} not found in {sheet_name}.{column_letter}")

        except KeyError:
            self.result_display.append(f"Sheet {sheet_name} not found.")
        except Exception as e:
            self.result_display.append(str(e))

    # Function to search for a target and print the entire row
    def search_in_workbook(self):
        sheet_name = self.sheet_input.text()
        column_letter = self.column_input.text()
        target =self.target_input.text()

        try:
            sheet = book[sheet_name]
            column = sheet[column_letter]

            for cell in column:
                if str(cell.value).casefold().strip() == str(target.casefold()).strip() or cell.value == target:
                    result = f'Found {target} in cell {column_letter}{cell.row}\n'
                    self.result_display.append(result)
                    
                    row = sheet[cell.row]
                    row_data = "\t".join([str(cell.value) for cell in row])
                    self.result_display.append(row_data)
                    return
            self.result_display.append(f"{target} not found in {sheet_name}.{column_letter}")

        except KeyError:
            self.result_display.append(f"Sheet {sheet_name} not found.")
        except Exception as e:
            self.result_display.append(str(e))

    # Function to list all tables in the selected sheet
    def list_tables(self):
        sheet_name = self.sheet_input.text()

        try:
            sheet = book[sheet_name]
            tables = sheet.tables  # Dictionary of tables in the sheet

            if tables:
                self.result_display.append(f"Tables in {sheet_name}:")
                for table_name, table in tables.items():
                    result = f'Table: {table_name}, Range: {table.ref}'
                    self.result_display.append(result)
            else:
                self.result_display.append(f"No tables found in {sheet_name}")

        except KeyError:
            self.result_display.append(f"Sheet {sheet_name} not found.")
        except Exception as e:
            self.result_display.append(str(e))


# Run the application
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelSearchApp()
    window.show()
    sys.exit(app.exec_())
