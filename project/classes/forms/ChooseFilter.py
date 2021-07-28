from PyQt5 import QtWidgets

from UIforms.ChooseFilter.ChooseFilter import Ui_Dialog_ChooseFilter

from classes.dicts import choose_position,choose_position_header,choose_position_header_evry_two

class ChooseFilter(QtWidgets.QDialog):

    def __init__(self, my_window):
        super(ChooseFilter, self).__init__()
        self.vivod_header = []
        self.my_window = my_window
        self.ui = Ui_Dialog_ChooseFilter()
        self.ui.setupUi(self)
        spisok = []
        self.vivod_dict = {}
        for key, value in choose_position.items():
            spisok.append(value)
        self.ui.listWidget.addItems(spisok)
        self.ui.listWidget.setSelectionMode(
            QtWidgets.QAbstractItemView.ExtendedSelection
        )

        self.ui.pushButton.clicked.connect(self.set_header_table2)
        self.ui.listWidget.verticalScrollBar().setStyleSheet('background: #444444')
        self.ui.listWidget.horizontalScrollBar().setStyleSheet('background: #444444')

    def set_header_table2(self):
        items = self.ui.listWidget.selectedItems()
        self.znach = []
        self.vivod_header = []
        for i in range(len(items)):
            self.znach.append(
                str(self.ui.listWidget.selectedItems()[i].text()))

        znach = self.znach
        for key, val in choose_position_header.items():
            for value in znach:
                if(value == val):
                    self.vivod_header.append(key)
                    self.vivod_header.append(choose_position_header_evry_two[key])
                
        self.my_window.ui.tableWidget_2.setColumnCount(len(self.vivod_header))
        self.my_window.ui.tableWidget_2.setHorizontalHeaderLabels(
            self.vivod_header)
        self.my_window.ui.tableWidget_2.resizeColumnsToContents()
        self.my_window.analizes.set_znach(self.znach)

        self.close()