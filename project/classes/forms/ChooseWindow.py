from PyQt5 import QtWidgets
from UIforms.ChooseForm.ChooseForm import Choose_Dialog

class ChooseWindow(QtWidgets.QDialog):
    # Класс окна для выбора листа в excel таблице
    def __init__(self, list):
        super(ChooseWindow, self).__init__()
        self.list = list
        self.ui = Choose_Dialog()
        self.ui.setupUi(self)
        self.ui.buttonBox.accepted.connect(self.ChooseList)
        self.ui.buttonBox.rejected.connect(self.RejectChooseList)
        self.setWindowTitle("Выбор листа для анализа")

        for value in self.list:
            self.ui.comboBox.addItem(str(value))

    def ChooseList(self):
        self.ChoosedSheet = self.ui.comboBox.currentText()

    def RejectChooseList(self):
        self.ChoosedSheet = 'None'
