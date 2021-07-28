from PyQt5 import QtWidgets
from classes.forms.MainForm import MyWindow

import sys

app = QtWidgets.QApplication([])
application = MyWindow()
application.setWindowTitle("Конвертер ведомости учета имущества")
application.show()

sys.exit(app.exec())
