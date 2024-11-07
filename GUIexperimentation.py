import sys

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import *

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.setWindowTitle("PyQt Demonstration")

        layout = QVBoxLayout()

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

        layout.addWidget(opener)
        layout.addWidget(closer)

        packer = QWidget()
        packer.setLayout(layout)
        self.setCentralWidget(packer)
