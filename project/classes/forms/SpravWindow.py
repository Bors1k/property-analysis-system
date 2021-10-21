from PyQt5 import QtWidgets, QtGui
from UIforms.SpravForm.SpravForm import Ui_MainWindow

class SpravWindow(QtWidgets.QMainWindow):

    def __init__(self):
        super(SpravWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowIcon(QtGui.QIcon(':roskazna.png'))
        self.setWindowTitle('Справочник')