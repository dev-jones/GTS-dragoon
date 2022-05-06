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
        print('나의 계좌번호는 : ' + account_number)
        return account_number
    
    # 종목 코드 가져오기
    def get_code_list_by_market(self, market_type):
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market_type)
        code_list = code_list.split(';')[:-1]
        return code_list
    
    # 종목명 가져오기
    def get_master_code_name(self, code): # 종목코드를 받아 종목명을 반환하는 함수
        code_name = self.dynamicCall('GetMasterCodeName(QString)', code)
        return code_name