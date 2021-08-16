from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QHeaderView

from UIforms.ChooseFilter.ChooseFilter import Ui_Dialog_ChooseFilter
from classes.otdel import Otdel

class ChooseOtdelFilter(QtWidgets.QDialog):
    # Класс окна для выбора отделов для анализа
    def __init__(self, my_window):
        super(ChooseOtdelFilter, self).__init__()
        self.my_window = my_window
        self.ui = Ui_Dialog_ChooseFilter()
        self.ui.setupUi(self)

        self.ui.listWidget.setSelectionMode(
            QtWidgets.QAbstractItemView.ExtendedSelection
        )
        self.ui.pushButton.clicked.connect(self.set_header_table2)
        self.ui.listWidget.verticalScrollBar().setStyleSheet('background: #444444')
        self.ui.listWidget.horizontalScrollBar().setStyleSheet('background: #444444')

    @QtCore.pyqtSlot(set)
    def getMnozh(self, set):
        self.mnozh = set
        self.sorted_list = list(self.mnozh)
        self.sorted_list.sort()
        self.ui.listWidget.clear()
        self.ui.listWidget.addItems(self.sorted_list)

    def set_header_table2(self):
        items = self.ui.listWidget.selectedItems()
        self.otdel = []
        self.my_window.otdels = []
        for i in range(len(items)):
            self.otdel.append(
                str(self.ui.listWidget.selectedItems()[i].text()))
            self.my_window.otdels.append(
                Otdel(self.ui.listWidget.selectedItems()[i].text()))

        self.my_window.ui.tableWidget_2.setRowCount(len(self.otdel))
        self.my_window.ui.tableWidget_2.setVerticalHeaderLabels(self.otdel)
        self.my_window.ui.tableWidget_2.resizeColumnsToContents()
        self.my_window.ui.tableWidget_2.verticalHeader(
        ).setSectionResizeMode(QHeaderView.Stretch)

        self.my_window.analizes.set_otdel(otdel=self.my_window.otdels)
        self.close()
