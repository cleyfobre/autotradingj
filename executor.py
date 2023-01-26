import sys
from app.kiwoom2 import *
from PyQt5.QtWidgets import *

class Executor:
    def __init__(self):

        # create new application object in memory
        self.app = QApplication(sys.argv)

        # run Kiwoom class for auto trading
        self.kiwoom = Kiwoom()

        # keep executing application until a time set on Kiwoom class
        self.app.exec_()
