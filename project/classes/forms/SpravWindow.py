from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import QHeaderView, QTableWidgetItem, QAbstractItemView
from UIforms.SpravForm.SpravForm import Ui_MainWindow
import json
import os


class SpravWindow(QtWidgets.QMainWindow):

    def __init__(self):
        super(SpravWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setMinimumWidth(510)
        self.setWindowIcon(QtGui.QIcon(':roskazna.png'))
        self.setWindowTitle('Справочник')
        self.ui.tableNorm.itemSelectionChanged.connect(self.on_selection)
        self.ui.tableWidget.itemSelectionChanged.connect(self.on_selection_table_two)
        self.ui.pushButton_3.clicked.connect(self.add_word_except)
        self.ui.pushButton_4.clicked.connect(self.del_word_except)
        self.words_of_exception = []
        self.new_position = True
        self.select_cell = ''
        self.ui.tableWidget.horizontalHeader().setVisible(False)
        self.ui.tableWidget.verticalHeader().setVisible(False)
        self.ui.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.ui.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.ui.tableWidget.verticalScrollBar().setStyleSheet('background: #444444')
        self.ui.tableWidget.horizontalScrollBar().setStyleSheet('background: #444444')
        self.ui.centralwidget.setStyleSheet('.QTableWidget .QTableCornerButton::section {background-color: rgba(0,0,0,0);')

    def on_selection(self):
        row = self.ui.tableNorm.currentItem().row()
        self.select_cell = self.ui.tableNorm.item(row, 0).text()
        self.table_zap()

    def on_selection_table_two(self):
        row_table_two = self.ui.tableWidget.currentItem().row()
        self.select_cell_table_two = self.ui.tableWidget.item(row_table_two, 0).text()


    def table_zap(self):
        self.ui.tableWidget.clear()
        self.ui.tableWidget.setRowCount(0)
        self.ui.tableWidget.setColumnCount(0)
        self.check_file = os.path.exists("C:\\Users\\Public\\property-analysis-system\\iskl.json")
        if self.check_file: 
            with open ("C:\\Users\\Public\\property-analysis-system\\iskl.json", encoding='utf-8') as f:
                    templates = json.load(f)

            for position in templates:
                if self.select_cell == position['position']:
                    self.ui.tableWidget.setRowCount(len(position['isckluchene']))
                    self.ui.tableWidget.setColumnCount(1)
                    schet = 0   
                    for isckluchene in position['isckluchene']:
                        self.ui.tableWidget.setItem(schet, 0, QTableWidgetItem(str(isckluchene)))
                        schet = schet + 1

    def del_word_except(self):
        self.check_file = os.path.exists("C:\\Users\\Public\\property-analysis-system\\iskl.json")
        if self.check_file:
            with open ("C:\\Users\\Public\\property-analysis-system\\iskl.json", encoding='utf-8') as f:
                templates = json.load(f)
                for dict_for_obj in templates:
                    obj = ObjectIckl(dict_for_obj['position'])
                    for iskl in dict_for_obj['isckluchene']:
                        obj.add_iskluchene(iskl)                
                    self.words_of_exception.append(obj)

        for position in self.words_of_exception:
            if position.position == self.select_cell:
                position.del_iskluchene(self.select_cell_table_two)

        with open("C:\\Users\\Public\\property-analysis-system\\iskl.json", "w",  encoding='utf-8') as outfile:
            json.dump(([ob.__dict__ for ob in self.words_of_exception]), outfile)

        self.table_zap()

    def add_word_except(self):
        self.words_of_exception.clear()
        self.check_file = os.path.exists("C:\\Users\\Public\\property-analysis-system\\iskl.json")
        if self.check_file:
            with open ("C:\\Users\\Public\\property-analysis-system\\iskl.json", encoding='utf-8') as f:
                templates = json.load(f)
                for dict_for_obj in templates:
                    obj = ObjectIckl(dict_for_obj['position'])
                    for iskl in dict_for_obj['isckluchene']:
                        obj.add_iskluchene(iskl)                
                    self.words_of_exception.append(obj)

        value_dict = self.ui.lineEdit.text()

        for position in self.words_of_exception:
            if position.position == self.select_cell:
                self.new_position = False
                position.add_iskluchene(value_dict)
        
        if(self.new_position):
            obj = ObjectIckl(self.select_cell)
            obj.add_iskluchene(value_dict)
            self.words_of_exception.append(obj)

        for position in self.words_of_exception:
            print(position.position)
            print(position.isckluchene)

        self.new_position = True
        self.check_file = os.path.exists("C:\\Users\\Public\\property-analysis-system\\iskl.json")
        if self.check_file:
            with open('C:\\Users\\Public\\property-analysis-system\\iskl.json', 'r+',  encoding='utf-8') as f:
                json.dump(([ob.__dict__ for ob in self.words_of_exception]), f)
        else:
            with open("C:\\Users\\Public\\property-analysis-system\\iskl.json", "w", encoding='utf-8') as outfile:
                json.dump(([ob.__dict__ for ob in self.words_of_exception]), outfile)
                self.check_file = False

        self.table_zap()

class ObjectIckl():
    def __init__(self, position):
        self.position = position
        self.isckluchene = []

        
    def add_iskluchene(self, value_dict):
        self.isckluchene.append(value_dict)

    def del_iskluchene(self, value_dict):
        self.isckluchene.remove(value_dict)

        
