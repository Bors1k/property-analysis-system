from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog
import openpyxl
from openpyxl import Workbook

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
        sheets = wb.sheetnames
        sheet = sheets[2]
        name_sheet = wb[sheet]

        # for row in range(9, name_sheet.max_row + 1):
        #     for column in "BCVJP":
        #         cell_name = "{}{}".format(column, row)
        #         print(name_sheet[cell_name].value)
        n = 1
        wb = Workbook()
        ws = wb.active
        for cell in name_sheet['A']:
            # print(cell.value)
            ws['A' + str(n)] = cell.value
            n = n + 1
        # wb.save('balances.xlsx')
        print('finished1')
        n = 1
        for cell in name_sheet['C']:
            # print(cell.value)
            ws['C' + str(n)] = cell.value
            n = n + 1
        wb.save('balances.xlsx')
        print('finished2')


app = QtWidgets.QApplication([])
application = mywindow()
application.show()

sys.exit(app.exec())
