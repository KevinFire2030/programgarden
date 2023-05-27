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

        ######### 초기 셋팅 함수들 바로 실행
        self.get_ocx_instance()  # OCX 방식을 파이썬에 사용할 수 있게 변환해 주는 함수
        self.event_slots()  # 키움과 연결하기 위한 시그널 / 슬롯 모음
        self.signal_login_commConnect()  # 로그인 요청 시그널 포함

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")  # 레지스트리에 저장된 api 모듈 불러오기


    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)  # 로그인 관련 이벤트


    def login_slot(self, err_code):
        # self.logging.logger.debug(errors(err_code)[1])
        print(errors(err_code))
        # 로그인 처리가 완료됐으면 이벤트 루프를 종료한다.
        self.login_event_loop.exit()

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")  # 로그인 요청 시그널

        self.login_event_loop.exec_()  # 이벤트루프 실행


