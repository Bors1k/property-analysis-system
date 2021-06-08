import xlwings
import xlwings as xw
import os


class Analyze:
    def __init__(self):
        self.dict_srok_isp = {}
        self.dict_name_otdel = {}
        self.dict_imushestvo = {}
        self.dict_kolvo = {}
        self.dict_srok_previshenia = {}

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
            self.dict_name_otdel[k] = value
            k = k + 1
        # Получаем все значения столбца B и записываем их в словарь
        result_znach = sh.range('B2:' + 'B' + str(rownum)).value
        k = 1
        for value in result_znach:
            self.dict_imushestvo[k] = value
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
            self.dict_srok_previshenia[k] = value
            k = k + 1
        for key, value in self.dict_srok_previshenia.items():
            print(str(key) + ' ' + str(value))

        wbxl.close()

        path = os.path.join(os.path.abspath(os.path.dirname(__file__)), filename)
        os.remove(path)
