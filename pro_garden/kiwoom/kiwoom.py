from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import *
from pro_garden.config.errorCode import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print("kiwoom() class start. ")

        ####### event loop를 실행하기 위한 변수모음
        self.login_event_loop = QEventLoop()  # 로그인 요청용 이벤트루프
        self.detail_account_info_event_loop = QEventLoop()

        #########################################

        ####### 변수모음
        self.account_num = None
        self.use_money = 0
        self.use_money_percent = 0.5
        self.account_stock_dict = {}

        #########################################

        ####### 스크린 번호 모음
        self.screen_my_info = "2000"

        #########################################

        ######### 초기 셋팅 함수들 바로 실행
        self.get_ocx_instance()  # OCX 방식을 파이썬에 사용할 수 있게 변환해 주는 함수
        self.event_slots()  # 키움과 연결하기 위한 시그널 / 슬롯 모음
        self.signal_login_commConnect()  # 로그인 요청 시그널 포함
        self.get_account_info()
        self.detail_account_info()
        self.detail_account_mystock()
        #########################################

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")  # 레지스트리에 저장된 api 모듈 불러오기

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)  # 로그인 관련 이벤트
        self.OnReceiveTrData.connect(self.trdata_slot)  # TR 관련 이벤트

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")  # 로그인 요청 시그널

        self.login_event_loop.exec_()  # 이벤트루프 실행

    def login_slot(self, err_code):
        # self.logging.logger.debug(errors(err_code)[1])
        print(errors(err_code))
        # 로그인 처리가 완료됐으면 이벤트 루프를 종료한다.
        self.login_event_loop.exit()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(QString)", "ACCNO")
        account_num = account_list.split(';')[0]
        self.account_num = account_num

        # self.logging.logger.debug("계좌번호 : %s" % account_num)
        print("계좌번호 : %s" % account_num)

    def detail_account_info(self):
        print("예수금을 요청하는 부분")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "빌밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "예수금상세현황요청", "opw00001", "0", self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):
        print("계좌평가잔고내역을 요청하는 부분")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "빌밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "계좌평가잔고내역요청", "opw00018", sPrevNext, self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

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
            print("예수금 %s" % deposit)
            print("예수금 형변환 %s" % int(deposit))

            self.use_money = int(deposit) * self.use_money_percent
            self.use_money = self.use_money / 4

            sub_price = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, 0, "권리대용금")
            print("권리대용금 %s" % sub_price)

            self.detail_account_info_event_loop.exit()

        if sRQName == "계좌평가잔고내역요청":
            total_buy_money = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, 0, "총매입금액")
            total_buy_money_result = int(total_buy_money)
            print("총 매입금액은 %s" % total_buy_money_result)

            total_profit_loss_rate = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, 0, "총수익률(%)")
            total_profit_loss_rate_result = float(total_profit_loss_rate)
            print("총 수익율은 %s" % total_profit_loss_rate_result)

            # 보유종목 가져오기 (45강 내용) : 20-10-07
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            cnt = 0

            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, i, "종목번호")
                code = code.strip()[1:] # 앞글자 영문 삭제

                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유수량")
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "수익률(%)")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                total_chegual_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입금액")
                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매매가능수량")

                if code in self.account_stock_dict:
                    pass
                else:
                    # self.account_stock_dict.update(code : {})
                    self.account_stock_dict[code] = {}

                code_nm = code_nm.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                learn_rate = float(learn_rate.strip())
                current_price = int(current_price.strip())
                total_chegual_price = int(total_chegual_price.strip())
                possible_quantity = int(possible_quantity.strip())

                self.account_stock_dict[code].update({"종목명": code_nm})
                self.account_stock_dict[code].update({"보유수량": stock_quantity})
                self.account_stock_dict[code].update({"매입가": buy_price})
                self.account_stock_dict[code].update({"수익률(%)": learn_rate})
                self.account_stock_dict[code].update({"현재가": current_price})
                self.account_stock_dict[code].update({"매입금액": total_chegual_price})
                self.account_stock_dict[code].update({"매매가능수량": possible_quantity})

                cnt = 1

            print("가지고 있는 종목 %s" % cnt)

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()
