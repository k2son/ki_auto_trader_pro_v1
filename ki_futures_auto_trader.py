import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop

class Strategy:
    def __init__(self, account, code, ocx):
        self.account = account
        self.code = code
        self.ocx = ocx
        self.data = []
        self.position = None
        self.entry_price = 0
        self.qty = 0

    def on_tick(self, price):
        if not self.data or self.data[-1]['time'] != pd.Timestamp.now().floor('1min'):
            self.data.append({'time': pd.Timestamp.now().floor('1min'),
                              'open': price, 'high': price, 'low': price, 'close': price})
        else:
            bar = self.data[-1]
            bar['high'] = max(bar['high'], price)
            bar['low'] = min(bar['low'], price)
            bar['close'] = price

        if len(self.data) >= 14:
            signal = self.check_entry_signal()
            if signal == 'long' and self.position != 'long':
                self.send_order(1)
            elif signal == 'short' and self.position != 'short':
                self.send_order(2)

            self.check_stop_loss(price)

    def check_entry_signal(self):
        df = pd.DataFrame(self.data[-14:])
        hh = df['high'].max()
        ll = df['low'].min()
        close = df['close'].iloc[-1]
        wr = (hh - close) / (hh - ll) * -100
        print(f"\u2193 WR: {wr:.2f}")
        if wr < -80:
            return 'long'
        elif wr > -20:
            return 'short'
        return None

    def check_stop_loss(self, price):
        if self.position == 'long':
            loss = (price - self.entry_price) / self.entry_price * 100
            if loss <= -18:
                print("\ud83d\udca5 롱 포지션 손절")
                self.send_order(2)
                self.position = None
        elif self.position == 'short':
            loss = (self.entry_price - price) / self.entry_price * 100
            if loss <= -15:
                print("\ud83d\udca5 숏 포지션 손절")
                self.send_order(1)
                self.position = None

    def send_order(self, order_type):
        self.ocx.dynamicCall(
            "SendOrder(QString, QString, QString, int, QString, int, QString, QString, QString)",
            ["주문", "0101", self.account, order_type, self.code, 1, "0", "03", ""]
        )
        print("\ud83d\udcc8 주문 전송 완료")

    def update_position(self, code, price, qty):
        self.entry_price = price
        self.qty = qty
        self.position = 'long' if qty > 0 else 'short'
        print(f"\ud83d\udccc 포지션 갱신: {self.position.upper()}, 진입가={price}, 수량={qty}")


class KiwoomApp:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.ocx = QAxWidget("KHOpenAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect[int].connect(self.on_login)
        self.ocx.OnReceiveRealData[str, str, str].connect(self.on_real_data)
        self.ocx.OnReceiveChejanData.connect(self.on_chejan)

        self.account = None
        self.code = "101S3000"
        self.strategy = None

        self.login_loop = QEventLoop()
        self.ocx.dynamicCall("CommConnect()")
        self.login_loop.exec_()

    def on_login(self, err_code):
        if err_code == 0:
            print("\u2705 로그인 성공")
            self.account = self.ocx.dynamicCall("GetLoginInfo(QString)", "ACCNO").split(';')[0]
            print("계좌번호:", self.account)

            self.strategy = Strategy(account=self.account, code=self.code, ocx=self.ocx)

            self.ocx.dynamicCall("SetRealReg(QString, QString, QString, QString)",
                                 ["1001", self.code, "10", "1"])
        else:
            print("\u274c 로그인 실패")
        self.login_loop.quit()

    def on_real_data(self, code, real_type, real_data):
        if code != self.code or real_type != "주식체결":
            return
        price_str = real_data.split('\t')[9]
        price = int(price_str)
        print(f"실시간 체결가: {price}")
        self.strategy.on_tick(price)

    def on_chejan(self, gubun, item_cnt, fid_list):
        if gubun == 1:
            order_no = self.ocx.dynamicCall("GetChejanData(int)", 9203).strip()
            code = self.ocx.dynamicCall("GetChejanData(int)", 9001).strip()[1:]
            name = self.ocx.dynamicCall("GetChejanData(int)", 302).strip()
            exec_price = self.ocx.dynamicCall("GetChejanData(int)", 910).strip()
            exec_qty = self.ocx.dynamicCall("GetChejanData(int)", 911).strip()

            print(f"\u2705 체결됨: 종목={code}, 이름={name}, 가격={exec_price}, 수량={exec_qty}")
            self.strategy.update_position(code=code, price=int(exec_price), qty=int(exec_qty))

    def execute(self):
        self.app.exec_()


if __name__ == "__main__":
    app = KiwoomApp()
    app.execute()
