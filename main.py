from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog
import openpyxl

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
        filename = QFileDialog.getOpenFileName(None, 'Открыть', os.path.dirname("C:\\"), 'All Files(*.xlsx)')
        wb = openpyxl.load_workbook(filename[0])
        # print(wb.get_sheet_names()[2])
        sheet = (wb.get_sheet_by_name(wb.get_sheet_names()[2]))
        massive = []
        for x in range(1,9,1):
            print(massive)
            massive.append([])
            print(massive[x-1])
            for y in range(1,24,1):
                massive[x-1].append(sheet.cell(row=x,column=y).value)

        print(massive)



app = QtWidgets.QApplication([])
application = mywindow()
application.show()

sys.exit(app.exec())
