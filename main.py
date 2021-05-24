from PyQt5 import QtWidgets, QtGui
from PyQt5.QtGui import QMovie
from PyQt5.QtWidgets import QFileDialog, QLabel, QTableWidgetItem, QAbstractItemView
import openpyxl
from openpyxl import Workbook
import re
from PyQt5.QtCore import QSize, QThread

from openpyxl.styles import Font, Alignment

from form import Ui_MainWindow  # импорт нашего сгенерированного файла
import sys
import os


class MyThread(QThread):
    def __init__(self, my_window, parent=None):
        super(MyThread, self).__init__()
        self.my_window = my_window

    def run(self):
        self.my_window.ui.pushButton_2.setEnabled(False)
        self.my_window.ui.pushButton_3.setEnabled(False)
        wb = openpyxl.load_workbook(self.my_window.filename[0])
        sheets = wb.sheetnames
        sheet = sheets[2]
        name_sheet = wb[sheet]
        n = 1
        k = 1
        self.my_window.wb = Workbook()
        ws = self.my_window.wb.active
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
        # wb.save('balances.xlsx')
        for cell in ws['B']:
            ws['F' + str(k)] = "=YEARFRAC(E" + str(k) + ",TODAY(),1)"
            ws['F' + str(k)].font = Font(size="8", name='Arial')
            ws['F' + str(k)].alignment = Alignment(horizontal='center', vertical='center')
            k = k + 1
            n = n + 1
        print('finished5')
        n = 1
        k = 1
        # wb.save('balances.xlsx')
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

        self.my_window.ui.tableWidget.setColumnCount(ws.max_column)
        self.my_window.ui.tableWidget.setHorizontalHeaderLabels(
            [str(ws['A1'].value), str(ws['B1'].value), str(ws['C1'].value),
             str(ws['D1'].value), str(ws['E1'].value), str(ws['F1'].value),
             str(ws['G1'].value), str(ws['H1'].value)])

        self.my_window.ui.tableWidget.setRowCount(ws.max_row)

        schet = 0
        for cell in ws['A']:
            self.my_window.ui.tableWidget.setItem(schet, 0, QTableWidgetItem(str(cell.value)))
            schet = schet + 1
        schet = 0
        for cell in ws['B']:
            self.my_window.ui.tableWidget.setItem(schet, 1, QTableWidgetItem(str(cell.value)))
            schet = schet + 1
        schet = 0
        for cell in ws['C']:
            self.my_window.ui.tableWidget.setItem(schet, 2, QTableWidgetItem(str(cell.value)))
            schet = schet + 1
        schet = 0
        for cell in ws['D']:
            self.my_window.ui.tableWidget.setItem(schet, 3, QTableWidgetItem(str(cell.value)))
            schet = schet + 1
        schet = 0
        for cell in ws['E']:
            self.my_window.ui.tableWidget.setItem(schet, 4, QTableWidgetItem(str(cell.value)))
            schet = schet + 1
        schet = 0
        for cell in ws['F']:
            self.my_window.ui.tableWidget.setItem(schet, 5, QTableWidgetItem(str(cell.value)))
            schet = schet + 1
        schet = 0
        for cell in ws['G']:
            self.my_window.ui.tableWidget.setItem(schet, 6, QTableWidgetItem(str(cell.value)))
            schet = schet + 1
        schet = 0
        for cell in ws['H']:
            self.my_window.ui.tableWidget.setItem(schet, 7, QTableWidgetItem(str(cell.value)))
            schet = schet + 1
        self.my_window.ui.tableWidget.resizeColumnsToContents()
        # for cell in ws['D']:
        #     cell.number_format = '0'
        ws.auto_filter.ref = ws.dimensions
        # wb.save('balances.xlsx')
        print('finishedEnd')
        self.my_window.ui.pushButton_2.setEnabled(True)
        self.my_window.ui.pushButton_3.setEnabled(True)
        self.my_window.movie.stop()
        self.my_window.ui.label_animation.setMovie(None)


class MyWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.setWindowIcon(QtGui.QIcon('roskazna.png'))
        self.ui.label_animation = QLabel(self)
        self.ui.label_animation.setFixedHeight(227)
        self.ui.label_animation.setFixedWidth(128)
        self.ui.label_animation.move(int(self.width() * 0.5) - int(self.ui.label_animation.width() * 0.5),
                                     int(self.height() * 0.5) - int(self.ui.label_animation.height() * 0.5))
        self.ui.pushButton_2.clicked.connect(self.btn_clicked)
        self.ui.pushButton_3.clicked.connect(self.save_btn_clicked)
        self.ui.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.ui.tableWidget.horizontalHeader().setStretchLastSection(True)
        stylesheet = "::section{background-color:rgb(252, 246, 5);}"
        self.ui.tableWidget.setStyleSheet('.QTableCornerButton::section{background-color: rgba(143, 144, 146, 100);}')
        self.ui.tableWidget.horizontalHeader().setStyleSheet(stylesheet)
        self.ui.tableWidget.verticalHeader().setStyleSheet(stylesheet)
        self.setStyleSheet('.QWidget {border-image: url(1C.png);}')
        self.ui.tableWidget.setStyleSheet('.QTableWidget {background-color: rgba(143, 144, 146, 100);border-radius: '
                                          '8px;}')
        self.ui.pushButton_2.setStyleSheet('.QPushButton{background-color: rgba(143, 144, 146, 80);} '
                                           '.QPushButton:hover{background-color: rgba(143, 144, 146, 130);}')
        self.ui.pushButton_3.setStyleSheet('.QPushButton{background-color: rgba(143, 144, 146, 80);} '
                                           '.QPushButton:hover{background-color: rgba(143, 144, 146, 130);}')
        self.ui.menubar.setStyleSheet('.QMenuBar{background-color: #444444;color: white;}')
        self.ui.statusbar.setStyleSheet('.QStatusBar{background-color: #444444;color: white;}')
        self.ui.tableWidget.verticalScrollBar().setStyleSheet('background: #444444')
        self.ui.tableWidget.horizontalScrollBar().setStyleSheet('background: #444444')

        self.my_thread = MyThread(my_window=self)

    def resizeEvent(self, event):
        self.ui.label_animation.move(int(self.width() * 0.5) - int(self.ui.label_animation.width() * 0.5),
                                     int(self.height() * 0.5) - int(self.ui.label_animation.height() * 0.5))
        QtWidgets.QMainWindow.resizeEvent(self, event)

    def new_thread(self):
        self.my_thread.start()
        # self.movie = QMovie('Ajax-loader.gif')
        self.movie = QMovie('Spinner.gif')
        self.movie.setScaledSize(QSize(128, 227))
        self.ui.label_animation.setMovie(self.movie)
        self.movie.start()

    def btn_clicked(self):
        self.filename = QFileDialog.getOpenFileName(None, 'Открыть', os.path.dirname("C:\\"), 'All Files(*.xlsx)')
        self.new_thread()

    def save_btn_clicked(self):
        file_save, _ = QFileDialog.getSaveFileName(self, 'Сохранить', 'Сводный перечень имущества', 'All Files(*.xlsx)')
        self.wb.save(file_save)


app = QtWidgets.QApplication([])
application = MyWindow()
application.setWindowTitle("Конвертер ведомости учета имущества")
application.show()

sys.exit(app.exec())
