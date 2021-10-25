from PyQt5 import QtWidgets
import datetime
import json
from classes.forms.SpravWindow import ObjectIckl
from UIforms.ChooseFilter.ChooseFilter import Ui_Dialog_ChooseFilter
from classes.dicts import Dictionary


class ChooseFilter(QtWidgets.QDialog):
    def __init__(self, my_window):
        super(ChooseFilter, self).__init__()
        self.vivod_header = []
        self.my_window = my_window
        self.ui = Ui_Dialog_ChooseFilter()
        self.ui.setupUi(self)
        spisok = []
        self.vivod_dict = {}
        self.dictionary = Dictionary()
        self.dictionary.zapoln_dict()
        for key, value in self.dictionary.choose_position.items():
            spisok.append(value)
        self.words_of_exception = []
        with open ("C:\\Users\\Public\\property-analysis-system\\iskl.json", encoding='utf-8') as f:
            templates = json.load(f)
            for dict_for_obj in templates:
                obj = ObjectIckl(dict_for_obj['position'])
                for iskl in dict_for_obj['isckluchene']:
                    obj.add_iskluchene(iskl)                
                self.words_of_exception.append(obj)

        for obj in self.words_of_exception:
            for isckluchene in obj.isckluchene:
                try:
                    spisok.remove(isckluchene.lower())
                except ValueError:
                    print('не удалось удалить ' + isckluchene)
        
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
        for key, val in self.dictionary.choose_position_header.items():
            for value in znach:
                if(value == val):
                    self.vivod_header.append(key)
                    self.vivod_header.append(self.dictionary.choose_position_header_evry_two[key])
                    self.vivod_header.append(self.dictionary.choose_position_header_evry_two[key] + " в " + str(self.my_window.ui.comboBox.currentText())  + " году")
                
        self.my_window.ui.tableWidget_2.setColumnCount(len(self.vivod_header))
        self.my_window.ui.tableWidget_2.setHorizontalHeaderLabels(
            self.vivod_header)
        self.my_window.ui.tableWidget_2.resizeColumnsToContents()
        self.my_window.analizes.set_znach(self.znach)

        self.close()