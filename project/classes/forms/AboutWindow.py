from PyQt5 import QtWidgets
from UIforms.AboutForm.AboutForm import Ui_Dialog

class AboutWindows(QtWidgets.QDialog):

    def __init__(self):
        super(AboutWindows, self).__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.setWindowTitle('О программе')