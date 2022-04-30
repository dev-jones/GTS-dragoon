from PyQt5.QAxContainer import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self._make_kiwoom_instance()
        self._set_signal_slots()
        self._comm_connect()
        
        self.account_number = self.get_account_number()
        
    def _make_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")
        
    def _set_signal_slots(self):
        self.OnEventConnect.connect(self._login_slot)
        
    def _login_slot(self, err_code):
        if err_code == 0:
            print('connected')
        else:
            print('not connected')
        
        self.login_event_loop.exit()
        
    def _comm_connect(self):                    # 로그인 함수: 로그인 요청 신호를 보낸 이후 응답 대기를 설정하는 함수
        self.dynamicCall('CommConnect()')
        
        self.login_event_loop = QEventLoop()    # 로그인 시도 결과에 대한 응답 대기 시작
        self.login_event_loop.exec_()
        
    def get_account_number(self, tag='ACCNO'):
        account_list = self.dynamicCall('GetLoginInfo(QString)', tag)
        account_number = account_list.split(';')[0]
        print(account_number)
        return account_number