from classes.shipment import Shipment
import xlwings as xw
# from classes.dicts import choose_position, choose_position_header
from classes.dicts import Dictionary
class Analyze:
    def __init__(self, my_window):
        self.my_window = my_window
        self.dict_srok_isp = {}
        self.dict_name_otdel = {}
        self.dict_imushestvo = {}
        self.dict_kolvo = {}
        self.dict_srok_previshenia = {}
        self.dict_expired_in_next_year = {}
        self.otdels = []
        self.znach = []

        self.select_otdel = {}
        self.select_imushestvo = {}

        self.result_dict_otdel_imushestvo = {}
        self.new_select_otdel = {}
        self.new_select_imushestvo = {}
        self.imush = ''
        self.dictionary = Dictionary()
        self.dictionary.zapoln_dict()

    def set_otdel(self, otdel):
        self.otdels = otdel

    def set_znach(self, znach):
        self.dictionary.zapoln_dict()
        self.znach = []
        for zn in znach:
            for key in self.dictionary.choose_position:
                if(zn==self.dictionary.choose_position[key]):
                    self.znach.append(str(key).lower())

    def analyze_xls(self, filename):
        self.dictionary.zapoln_dict()
        app = xw.App(visible=False)
        wbxl = xw.Book(filename)
        sh = wbxl.sheets[0]
        for item in self.otdels:
            item.getName()

        # Получаем максимальую ячейку
        self.rownum = sh.range('F1').end('down').last_cell.row

        # Получаем все значения столбца F и записываем их в словарь
        result_znach = sh.range('F2:' + 'F' + str(self.rownum)).value
        k = 2
        for value in result_znach:
            self.dict_srok_isp[k] = value
            k = k + 1
        # Получаем все значения столбца A и записываем их в словарь
        result_znach = sh.range('A2:' + 'A' + str(self.rownum)).value
        k = 2
        for value in result_znach:
            if value is None:
                self.dict_name_otdel[k] = " "
            else:
                self.dict_name_otdel[k] = str(value).lower()
            k = k + 1
        # Получаем все значения столбца B и записываем их в словарь
        result_znach = sh.range('B2:' + 'B' + str(self.rownum)).value
        k = 2
        for value in result_znach:
            self.dict_imushestvo[k] = str(value).lower()
            k = k + 1
        # Получаем все значения столбца C и записываем их в словарь
        result_znach = sh.range('C2:' + 'C' + str(self.rownum)).value
        k = 2
        for value in result_znach:
            self.dict_kolvo[k] = value
            k = k + 1
        # Получаем все значения столбца H и записываем их в словарь
        result_znach = sh.range('H2:' + 'H' + str(self.rownum)).value
        k = 2
        for value in result_znach:
            self.dict_srok_previshenia[k] = str(value).lower()
            k = k + 1

        # Получаем все значения столбца K и записываем их в словарь
        result_znach = sh.range('K2:' + 'K' + str(self.rownum)).value
        k = 2
        for value in result_znach:
            if(str(value).lower() == 'да'):
                self.dict_expired_in_next_year[k] = True
            else:
                self.dict_expired_in_next_year[k] = False
            k = k + 1


        wbxl.close()

        for item in self.otdels:
            for x in range(2,self.rownum):
                if(item.getName().lower() == self.dict_name_otdel[x]):
                    for value in self.znach:
                        if(value.lower() in self.dict_imushestvo[x]):
                            for key in self.dictionary.choose_position_header:
                                if(self.dictionary.choose_position_header[key]==self.dictionary.choose_position[value]):
                                    if 'стол' in self.dictionary.choose_position_header[key] and 'настол' in self.dict_imushestvo[x]:
                                        pass
                                    else:
                                        shipment = Shipment(key)
                                        item.addNewShipment(item.shipments,shipment,self.dict_kolvo[x],self.dict_srok_previshenia[x],self.dict_expired_in_next_year[x])

        return self.otdels

      