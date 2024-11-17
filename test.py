import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QComboBox, QLineEdit, QPushButton, QTextEdit
)
from openpyxl import load_workbook

class PetDataFilterApp(QWidget):
    def __init__(self):
        super().__init__()

        # Load Excel Workbook
        self.book = load_workbook("FoundHouse.xlsx")  # Update the path to the correct file

        # GUI Window Setup
        self.setWindowTitle("Pet Data Filter")
        self.setGeometry(100, 100, 600, 400)
        layout = QVBoxLayout()

        # Species selection
        self.species_label = QLabel("Select Species:")
        self.species_input = QComboBox()
        self.species_input.addItems(["Dog", "Cat", "Others"])  # Predefined options

        # Input fields for filtering
        self.time_label = QLabel("Enter Time (e.g., days):")
        self.time_input = QLineEdit()

        self.fixed_status_label = QLabel("Fixed Status (Yes/No):")
        self.fixed_status_input = QComboBox()
        self.fixed_status_input.addItems(["Yes", "No"])

        self.single_digit_label = QLabel("Enter Single-Digit Field:")
        self.single_digit_input = QLineEdit()

        # Buttons for filtering options (no actions connected)
        self.search_time_button = QPushButton("Search by Time Returned to Owner")
        self.compare_fixed_button = QPushButton("Compare Fixed vs. Non-Fixed")
        self.filter_single_digit_button = QPushButton("Filter by Single-Digit Field")

        # Output area
        self.result_display = QTextEdit()
        self.result_display.setReadOnly(True)

        # Add widgets to layout
        layout.addWidget(self.species_label)
        layout.addWidget(self.species_input)
        layout.addWidget(self.time_label)
        layout.addWidget(self.time_input)
        layout.addWidget(self.fixed_status_label)
        layout.addWidget(self.fixed_status_input)
        layout.addWidget(self.single_digit_label)
        layout.addWidget(self.single_digit_input)
        layout.addWidget(self.search_time_button)
        layout.addWidget(self.compare_fixed_button)
        layout.addWidget(self.filter_single_digit_button)
        layout.addWidget(self.result_display)

        self.setLayout(layout)

# Run the application
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PetDataFilterApp()
    window.show()
    sys.exit(app.exec_())
