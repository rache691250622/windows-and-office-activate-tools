import sys
import os
if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']
import sys
import time
import requests
import win32api
import win32con
import images_rc
from PyQt5.QtGui import QIcon
from operation import *
from PyQt5.QtWidgets import QMainWindow, QApplication
from mainwindow import Ui_MainWindow


def check_net():
    try:
        html = requests.get("http://www.baidu.com", timeout=2)
    except():
        win32api.MessageBox(0, "未连接网络，请联网后重试", "提醒", win32con.MB_ICONWARNING)
        return
    return True


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)
        self.setWindowIcon(QIcon(':\images\icon.png'))
        self.activate = None

    def get_windows(self):
        if not self.activate:
            self.activate = activate(self)
        self.activate.show_text.connect(self.knock)
        self.activate.start()

    def get_office(self):
        if not self.activate:
            self.activate = activate(self)
        self.activate.show_text.connect(self.knock)
        self.activate.start()

    def knock(self, islock):
        if islock == 'lock':
            self.windows_key_input.setEnabled(False)
            self.windows_server_input.setEnabled(False)
            self.windows_ver.setEnabled(False)
            self.office_key_input.setEnabled(False)
            self.office_server_input.setEnabled(False)
            self.office_ver.setEnabled(False)
            self.tabWidget.setEnabled(False)
        elif islock == 'block':
            self.windows_key_input.setEnabled(True)
            self.windows_server_input.setEnabled(True)
            self.windows_ver.setEnabled(True)
            self.office_key_input.setEnabled(True)
            self.office_server_input.setEnabled(True)
            self.office_ver.setEnabled(True)
            self.tabWidget.setEnabled(True)

    def input_clear(self):
        self.windows_ver.setCurrentIndex(-1)
        self.windows_key_input.clear()
        self.windows_server_input.clear()
        self.office_ver.setCurrentIndex(-1)
        self.office_key_input.clear()
        self.office_server_input.clear()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWindow()
    w.windows_go.clicked.connect(check_net)
    windows = w.windows_go.clicked.connect(w.get_windows)
    w.office_go.clicked.connect(check_net)
    w.office_go.clicked.connect(w.get_office)
    w.tabWidget.tabBarClicked.connect(w.input_clear)
    w.show()
    sys.exit(app.exec_())
