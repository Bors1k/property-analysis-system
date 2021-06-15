import xlwings
import xlwings as xw
import os
from PyQt5 import QtCore
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QTableWidgetItem


class Analyze:
    def __init__(self, my_window):
        self.my_window = my_window
        self.dict_srok_isp = {}
        self.dict_name_otdel = {}
        self.dict_imushestvo = {}
        self.dict_kolvo = {}
        self.dict_srok_previshenia = {}
        self.otdel = []
        self.znach = []

        self.select_otdel = {}
        self.select_imushestvo = {}

        self.result_dict_otdel_imushestvo = {}
        self.new_select_otdel = {}
        self.new_select_imushestvo = {}
        self.imush = ''

    def set_otdel(self, otdel):
        for otd in otdel:
            self.otdel.append(str(otd).lower())

        print(self.otdel)

    def set_znach(self, znach):
        for zn in znach:
            self.znach.append(str(zn).lower())

        print(self.znach)

    def analyze_xls(self, filename):
        app = xw.App(visible=False)
        wbxl = xw.Book(filename)
        sh = wbxl.sheets[0]
        # Получаем максимальую ячейку
        rownum = sh.range('F1').end('down').last_cell.row
        print(rownum)
        # Получаем все значения столбца F и записываем их в словарь
        result_znach = sh.range('F2:' + 'F' + str(rownum)).value
        k = 1
        for value in result_znach:
            self.dict_srok_isp[k] = value
            k = k + 1
        # Получаем все значения столбца A и записываем их в словарь
        result_znach = sh.range('A2:' + 'A' + str(rownum)).value
        k = 1
        for value in result_znach:
            if value is not None:
                self.dict_name_otdel[k] = str(value).lower()
            k = k + 1
        # Получаем все значения столбца B и записываем их в словарь
        result_znach = sh.range('B2:' + 'B' + str(rownum)).value
        k = 1
        for value in result_znach:
            self.dict_imushestvo[k] = str(value).lower()
            k = k + 1
        # Получаем все значения столбца C и записываем их в словарь
        result_znach = sh.range('C2:' + 'C' + str(rownum)).value
        k = 1
        for value in result_znach:
            self.dict_kolvo[k] = value
            k = k + 1
        # Получаем все значения столбца H и записываем их в словарь
        result_znach = sh.range('H2:' + 'H' + str(rownum)).value
        k = 1
        for value in result_znach:
            self.dict_srok_previshenia[k] = str(value).lower()
            k = k + 1
        # for key, value in self.dict_srok_previshenia.items():
        #     print(str(key) + ' ' + str(value))

        wbxl.close()

        # path = os.path.join(os.path.abspath(os.path.dirname(__file__)), filename)
        # os.remove(path)

    def calculate(self):
        self.otdel.sort()
        for name_otdel_key, name_otdel_value in self.dict_name_otdel.items():
            for otdel in self.otdel:
                if otdel in name_otdel_value:
                    self.select_otdel[name_otdel_key] = name_otdel_value

        for imushestvo_key, imushestvo_value in self.dict_imushestvo.items():
            for znach in self.znach:
                if znach in imushestvo_value:
                    self.select_imushestvo[imushestvo_key] = imushestvo_value


        for select_otdel_key, select_otdel_value in self.select_otdel.items():
            for select_imushestvo_key, select_imushestvo_value in self.select_imushestvo.items():
                if select_otdel_key == select_imushestvo_key:
                    for kolvo_key, kolvo_value in self.dict_kolvo.items():
                        for sel_znach in self.znach:
                            if sel_znach in select_imushestvo_value:
                                kolvo_result = 0
                                if select_imushestvo_key == kolvo_key:
                                    kolvo_result = kolvo_result + int(kolvo_value)
                                if kolvo_result != 0:
                                    # print(select_otdel_value + " " + sel_znach + " " + str(kolvo_result))
                                    self.new_select_otdel = {kolvo_key: select_otdel_value}
                                    self.new_select_imushestvo = {kolvo_key: sel_znach}
                                    self.new_select_kolvo = {kolvo_key: kolvo_result}

            # dict_im_val = {}
            # for new_select_otdel_key, new_select_otdel_value in self.new_select_otdel.items():
            #
            #     for new_select_imushestvo_key, new_select_imushestvo_value in self.new_select_imushestvo.items():
            #         if new_select_otdel_key == new_select_imushestvo_key:
            #             result_imushestvo = result_imushestvo + self.new_select_kolvo[new_select_imushestvo_key]
            #             dict_im_val = {new_select_imushestvo_value, str(result_imushestvo)}
            #             print(dict_im_val)
            #     result_imushestvo = 0
            print(str(self.new_select_otdel.items()) + ' ' + str(self.new_select_imushestvo.items()))
            result_imushestvo = 0
            k = 0
            for new_select_otdel_key, new_select_otdel_value in self.new_select_otdel.items():
                for new_select_imushestvo_key, new_select_imushestvo_value in self.new_select_imushestvo.items():
                    if new_select_otdel_key == new_select_imushestvo_key:
                        if k == 0:
                            imush = new_select_imushestvo_value
                            k = 1
                        # imush = Tumba huyumba
                        if imush in new_select_imushestvo_value:
                            result_imushestvo = result_imushestvo + self.new_select_kolvo[new_select_imushestvo_key]
                        # result_imushestvo = 0 + 3
                        else:
                            print(new_select_otdel_value + ' ' + str(result_imushestvo) + ' ' + new_select_imushestvo_value)
                            imush = new_select_imushestvo_value
                            # imush = Shkaf zolotoy
                            result_imushestvo = 0