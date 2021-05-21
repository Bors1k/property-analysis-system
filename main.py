from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog
import openpyxl
from openpyxl import Workbook
import re

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
        k = 1
        wb = Workbook()
        ws = wb.active
        znach = ''
        for cell in name_sheet['A']:
            if re.match(r'\d', str(cell.value)):
                if name_sheet['C' + str(n)].value is not None:
                #     ws['A' + str(n)] = None
                # else:
                    ws['A' + str(k)] = znach
                    k = k + 1
            else:
                if name_sheet['C' + str(n)].value is None:
                    # ws['A' + str(n)] = None
                    znach = cell.value
                else:
                    znach = cell.value
                    ws['A' + str(k)] = znach
                    k = k + 1


            # try:
            #     val = int(cell.value)
            #     if znach != '':
            #         ws['A' + str(n)] = znach
            # except:
            #     znach = cell.value
            #     ws['A' + str(n)] = znach
            # ws['A' + str(n)] = cell.value

            n = n + 1
        # wb.save('balances.xlsx')
        print('finished1')
        n = 1
        k = 1
        for cell in name_sheet['C']:
            if name_sheet['C' + str(n)].value is not None:
                ws['B' + str(k)] = cell.value
                k = k + 1

            # print(cell.value)
            # ws['B' + str(n)] = cell.value
            n = n + 1
        # wb.save('balances.xlsx')
        print('finished2')
        n = 1
        k = 1
        for cell in name_sheet['J']:
            if name_sheet['C' + str(n)].value is not None:
                ws['C' + str(k)] = cell.value
                k = k + 1
            # print(cell.value)
            # ws['C' + str(n)] = cell.value
            n = n + 1
        # wb.save('balances.xlsx')
        print('finished3')
        n = 1
        k = 1
        for cell in name_sheet['P']:
            if name_sheet['C' + str(n)].value is not None:
                ws['D' + str(k)] = cell.value
                k = k + 1
            # print(cell.value)
            # ws['D' + str(n)] = cell.value
            n = n + 1
        wb.save('balances.xlsx')
        print('finished4')


app = QtWidgets.QApplication([])
application = mywindow()
application.show()

sys.exit(app.exec())
