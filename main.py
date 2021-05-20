from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog

from form import Ui_MainWindow  # импорт нашего сгенерированного файла
import sys
import os


class mywindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(mywindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton_2.clicked.connect(self.btnclicked)

    def btnclicked(self):
        filename = QFileDialog.getOpenFileName(None, 'Открыть', os.path.dirname("C:\\"), 'All Files(*.xls)')
        print(filename)




app = QtWidgets.QApplication([])
application = mywindow()
application.show()

sys.exit(app.exec())
