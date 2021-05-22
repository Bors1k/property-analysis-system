from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog
import openpyxl
from openpyxl import Workbook
import re

from openpyxl.styles import Font, NamedStyle, Alignment

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
        sheets = wb.sheetnames
        sheet = sheets[2]
        name_sheet = wb[sheet]
        n = 1
        k = 1
        wb = Workbook()
        ws = wb.active
        ws.title = "Сводный перечень имущества"


        ws.row_dimensions[1].ht = 39.6
        znach = ''
        for cell in name_sheet['A']:
            if re.match(r'\d', str(cell.value)):
                if name_sheet['C' + str(n)].value is not None:
                    ws['A' + str(k)] = znach
                    ws['A' + str(k)].font = Font(size="8", name='Arial')
                    ws['A' + str(k)].alignment = Alignment(horizontal='center', vertical='center')
                    k = k + 1
            else:
                if name_sheet['C' + str(n)].value is None:
                    znach = cell.value
                else:
                    znach = cell.value
                    ws['A' + str(k)] = znach
                    ws['A' + str(k)].font = Font(size="8", name='Arial')
                    ws['A' + str(k)].alignment = Alignment(horizontal='center', vertical='center')
                    k = k + 1
            n = n + 1
        print('finished1')
        n = 1
        k = 1
        for cell in name_sheet['C']:
            if name_sheet['C' + str(n)].value is not None:
                ws['B' + str(k)] = cell.value
                ws['B' + str(k)].font = Font(size="8", name='Arial')
                ws['B' + str(k)].alignment = Alignment(horizontal='center', vertical='center')
                k = k + 1
            n = n + 1
        print('finished2')
        n = 1
        k = 1
        for cell in name_sheet['J']:
            if name_sheet['C' + str(n)].value is not None:
                ws['D' + str(k)] = cell.value
                ws['D' + str(k)].font = Font(size="8", name='Arial')
                ws['D' + str(k)].alignment = Alignment(horizontal='center', vertical='center')
                ws['D' + str(k)].number_format = '0'
                k = k + 1
            n = n + 1
        print('finished3')
        n = 1
        k = 1
        # my_date_style = NamedStyle(name='my_date_style', number_format='dd.mm.yyyy')
        for cell in name_sheet['P']:
            if name_sheet['C' + str(n)].value is not None:
                ws['E' + str(k)] = cell.value
                ws['E' + str(k)].font = Font(size="8", name='Arial')
                ws['E' + str(k)].alignment = Alignment(horizontal='center', vertical='center')
                # ws['E' + str(k)].style = my_date_style
                k = k + 1
            n = n + 1
        print('finished4')
        n = 1
        k = 1
        for cell in name_sheet['V']:
            if name_sheet['C' + str(n)].value is not None:
                ws['C' + str(k)] = cell.value
                ws['C' + str(k)].font = Font(size="8", name='Arial')
                ws['C' + str(k)].alignment = Alignment(horizontal='center', vertical='center')
                k = k + 1
            n = n + 1
        n = 1
        k = 1
        # ws.delete_rows(2, 1)
        wb.save('balances.xlsx')
        for cell in ws['B']:
            ws['F' + str(k)] = "=YEARFRAC(E" + str(k) + ",TODAY(),1)"
            ws['F' + str(k)].font = Font(size="8", name='Arial')
            ws['F' + str(k)].alignment = Alignment(horizontal='center', vertical='center')
            k = k + 1
            n = n + 1
        print('finished5')
        n = 1
        k = 1
        wb.save('balances.xlsx')
        for cell in ws['B']:
            ws['H' + str(k)] = "=IF((G" + str(k) + "-F" + str(k) + ")<0,F" + str(k) + "-G" + str(k) + ",\"в пределах " \
                                                                                                      "срока\") "
            ws['H' + str(k)].font = Font(size="8", name='Arial')
            ws['H' + str(k)].alignment = Alignment(horizontal='center', vertical='center')
            ws['H' + str(k)].number_format = '0.0'
            k = k + 1
            n = n + 1
        print('finished5')

        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 48
        ws.column_dimensions['C'].width = 7
        ws.column_dimensions['D'].width = 24
        ws.column_dimensions['E'].width = 14.5
        ws.column_dimensions['F'].width = 14.5
        ws.column_dimensions['H'].width = 24

        ws['A1'] = 'Отдел'
        ws['A1'].font = Font(bold=True, size="10", name='Arial')
        ws['A1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        ws['B1'] = 'Наименование имущества'
        ws['B1'].font = Font(bold=True, size="10", name='Arial')
        ws['B1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        ws['C1'] = 'Кол-во'
        ws['C1'].font = Font(bold=True, size="10", name='Arial')
        ws['C1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        ws['D1'] = 'Инвентарный номер'
        ws['D1'].font = Font(bold=True, size="10", name='Arial')
        ws['D1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        ws['E1'] = 'Дата принятия к учету'
        ws['E1'].font = Font(bold=True, size="10", name='Arial')
        ws['E1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        ws['F1'] = 'Срок использования'
        ws['F1'].font = Font(bold=True, size="10", name='Arial')
        ws['F1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        ws['H1'] = 'Срок превышен, лет'
        ws['H1'].font = Font(bold=True, size="10", name='Arial')
        ws['H1'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        # for cell in ws['D']:
        #     cell.number_format = '0'
        ws.auto_filter.ref = ws.dimensions
        wb.save('balances.xlsx')
        print('finishedEnd')


app = QtWidgets.QApplication([])
application = mywindow()
application.show()

sys.exit(app.exec())
