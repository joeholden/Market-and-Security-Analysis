from PyQt5.QtCore import QSize, Qt
from PyQt5.QtWidgets import QApplication, QWidget, QMainWindow, QPushButton
import sys


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Portfolio Analysis')
        self.setFixedSize(QSize(1600, 1200))


app = QApplication(sys.argv)
window = MainWindow()
window.show()

app.exec_()
