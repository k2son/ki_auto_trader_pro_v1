import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QAxContainer import QAxWidget

class KiwoomConnect:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.ocx = QAxWidget("KHOpenAPI.KHOpenAPICtrl.1")
        self.ocx.dynamicCall("CommConnect()")  # 로그인창 띄우기
        self.app.exec_()

if __name__ == "__main__":
    KiwoomConnect()
