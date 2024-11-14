import sys
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import *
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import pandas as pd

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.setWindowTitle("Found House - GUI Test")

        self.layout = QVBoxLayout()

        opener = QLabel("This is a quick test that I threw together exploring the PyQt GUI's possibilities.")
        font = opener.font()
        font.setPointSize(20)
        opener.setFont(font)
        opener.setAlignment(Qt.AlignTop | Qt.AlignHCenter)

        closer = QLabel("It's SUPER bare-bones, I know.")
        font = closer.font()
        font.setPointSize(12)
        closer.setFont(font)
        closer.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)

        self.layout.addWidget(opener)
        self.layout.addWidget(closer)

        self.load_data_button = QPushButton("Load Excel File")
        self.load_data_button.clicked.connect(self.load_data)
        self.layout.addWidget(self.load_data_button)

        self.data_label = QLabel("No data loaded")
        self.layout.addWidget(self.data_label)

        packer = QWidget()
        packer.setLayout(self.layout)
        self.setCentralWidget(packer)

    def load_data(self):
        book = load_workbook("openpyxl/FoundHouse.xlsx")
        sheet = book.active
        print(book.sheetnames)
        data = []
        for row in sheet.rows:
            data.append([cell.value for cell in row])
        self.display_data(data)

    def display_data(self, data):
        text = ""
        for row in data:
            text += str(row) + "\n"
        self.data_label.setText(text)

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(2000,1000)
    window.move(100,100)
    window.show()
    app.exec()

if __name__ == '__main__':
    main()