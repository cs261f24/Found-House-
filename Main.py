import sys

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import *
import GUIexperimentation

def main():
    # code to connect Excel sheet to this program
    # loop creating all the necessary Pet and Owner objects
    app = QApplication(sys.argv)
    window = GUIexperimentation.MainWindow()
    window.resize(2000,1000)
    window.move(100,100)
    window.show()
    app.exec()

if __name__ == '__main__':
    main()
