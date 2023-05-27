import os

from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import *
from pro_garden.config.errorCode import *
from PyQt5.QtTest import *
from pro_garden.config.kiwoomType import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print("kiwoom() class start. ")

        ####### event loop를 실행하기 위한 변수모음
        self.login_event_loop = QEventLoop()  # 로그인 요청용 이벤트루프
        self.detail_account_info_event_loop = QEventLoop()

        ####### 변수모음
        self.account_num = None
        self.use_money = 0
        self.use_money_percent = 0.5
        self.deposit = 0 #예수금
        self.output_deposit = 0  # 출금가능 금액

        ####### 스크린 번호 모음
        self.screen_my_info = "2000"

        ######### 초기 셋팅 함수들 바로 실행
        self.get_ocx_instance()  # OCX 방식을 파이썬에 사용할 수 있게 변환해 주는 함수
        self.event_slots()  # 키움과 연결하기 위한 시그널 / 슬롯 모음
        self.signal_login_commConnect()  # 로그인 요청 시그널 포함
        self.get_account_info()
        self.detail_account_info()




    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")  # 레지스트리에 저장된 api 모듈 불러오기


    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)  # 로그인 관련 이벤트
        self.OnReceiveTrData.connect(self.trdata_slot)  # TR 관련 이벤트


    def login_slot(self, err_code):
        # self.logging.logger.debug(errors(err_code)[1])
        print(errors(err_code))
        # 로그인 처리가 완료됐으면 이벤트 루프를 종료한다.
        self.login_event_loop.exit()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        """
        TR 요청을 받는 부분
        :param sScrNo: 스크린 번호
        :param sRQName: 사용자 구분명
        :param sTrCode: 요청 ID, TR 코드
        :param sRecordName: 사용안함
        :param sPrevNext: 다음 페이지가 있는지? (연속조회 유무를 판단하는 값 0: 연속(추가조회)데이터 없음, 1:연속(추가조회) 데이터 있음)
        :return:
        """

        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, 0, "예수금")
            self.deposit = int(deposit)
            print("예수금 %s" % self.deposit)

            self.use_money = int(deposit) * self.use_money_percent
            self.use_money = self.use_money / 4

            output_deposit = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, 0, "출금가능금액")
            self.output_deposit = int(output_deposit)
            print("출금가능금액 %s" % self.output_deposit)

            self.detail_account_info_event_loop.exit()

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")  # 로그인 요청 시그널

        self.login_event_loop.exec_()  # 이벤트루프 실행

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(QString)", "ACCLIST")
        account_num = account_list.split(';')
        self.account_num = account_num

        print("계좌번호 : %s" % account_num)


        # self.logging.logger.debug("계좌번호 : %s" % account_num)
        # 계좌번호: [' 선옵 7014194731', '상시 8049229611', '대주? 8049229711', '']


    def detail_account_info(self):
        print("예수금을 요청하는 부분")
        #print(self.screen_my_info)
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num[1])
        self.dynamicCall("SetInputValue(QString, QString)", "빌밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "예수금상세현황요청", "opw00001", "0", self.screen_my_info)
        #self.dynamicCall("CommRqData(QString, QString, int, QString)", "예수금상세현황요청", "opw00001", "0","1000")

        self.detail_account_info_event_loop.exec_()



