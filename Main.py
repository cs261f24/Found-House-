import sys

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import *
import GUIexperimentation
import read
import test

def main():
    book = read.load_workbook("FoundHouse.xlsx")
    read.option_search()
    print("read done, moving onto test")
    test.search_in_workbook()
    print("test done, moving onto rows")
    import rows
    # loop creating all the necessary Pet and Owner objects
    app = QApplication(sys.argv)
    window = GUIexperimentation.MainWindow()
    window.resize(2000,1000)
    window.move(100,100)
    window.show()
    app.exec()
    print("should be done now")

if __name__ == '__main__':
    main()
