from PyQt6.QtWidgets import QApplication, QMainWindow
from chartplot.MainWindowExt import MainWindowExt

app = QApplication([])
mainwindow = QMainWindow()
myui = MainWindowExt()
myui.setupUi(mainwindow)
myui.showWindow()
app.exec()
