from PyQt5.QAxContainer import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import time
import pandas as pd

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self._make_kiwoom_instance()
        self._set_signal_slots()
        self._comm_connect()
        
        self.account_number = self.get_account_number()
        
        self.tr_event_loop = QEventLoop()   # tr요청에 대한 응답 대기를 위한 변수
        
    def _make_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")
        
    def _set_signal_slots(self):
        self.OnEventConnect.connect(self._login_slot)
        self.OnReceiveTrData.connect(self._on_receive_tr_data)  # TR의 응답 결과를 _on_receive_tr_data로 설정
        
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
    
    # 종목의 상장일부터 가장 최근 일자까지 일봉 정보를 가져오는 함수
    def get_price_data(self, code):
        self.dynamicCall('SetInputValue(QString, QString)', '종목코드', code)
        self.dynamicCall('SetInputValue(QString, QString)', '수정주가구분', 1)
        self.dynamicCall('CommRqData(QString, QString, int, QString)', 'opt10081_req', 'opt10081', 0, '0001')
        
        self.tr_event_loop.exec_()
        
        ohlcv = self.tr_data
        
        while self.has_next_tr_data:
            self.dynamicCall('SetInputValue(QString, QString)', '종목코드', code)
            self.dynamicCall('SetInputValue(QString, QString)', '수정주가구분', '1')
            self.dynamicCall('CommRqData(QString, QString, int, QString)', 'opt10081_req', 'opt10081', 2, '0001')
            self.tr_event_loop.exec_()
            
            for key, val in self.tr_data.items():
                ohlcv[key][-1:] = val
        
        df = pd.DataFrame(ohlcv, columns=['open', 'high', 'low', 'close', 'volume'], index=ohlcv['date'])
        
        return df[::-1]

    
    # tr 조회의 응답 결과를 얻어 오는 함수
    def _on_receive_tr_data(self, screen_no, rqname, trcode, record_name, next, unused1, unused2, unused3, unused4):
        print('[Kiwoom] _onreceive_tr_data is called {} / {} / {}'.format(screen_no, rqname, trcode))
        tr_data_cnt = self.dynamicCall('GetRepeatCnt(QString, QString)', trcode, rqname)
        
        if next == '2':
            self.has_next_tr_data = True
        else:
            self.has_next_tr_data = False
            
        if rqname == 'opt10081_req':
            ohlcv = {'date': [], 'open': [], 'high': [], 'low': [], 'close': [], 'volume': []}
            
            for i in range(tr_data_cnt):
                date = self.dynamicCall('GetCommData(QString, QString, int, QString', trcode, rqname, i, '일자')
                open = self.dynamicCall('GetCommData(QString, QString, int, QString', trcode, rqname, i, '시가')
                high = self.dynamicCall('GetCommData(QString, QString, int, QString', trcode, rqname, i, '고가')
                low = self.dynamicCall('GetCommData(QString, QString, int, QString', trcode, rqname, i, '저가')
                close = self.dynamicCall('GetCommData(QString, QString, int, QString', trcode, rqname, i, '현재가')
                volume = self.dynamicCall('GetCommData(QString, QString, int, QString', trcode, rqname, i, '거래량')
                
                ohlcv['date'].append(date.strip())
                ohlcv['open'].append(int(open))
                ohlcv['high'].append(int(high))
                ohlcv['low'].append(int(low))
                ohlcv['close'].append(int(close))
                ohlcv['volume'].append(int(volume))
            
            self.tr_data = ohlcv
    
        self.tr_event_loop.exit()
        time.sleep(0.5)