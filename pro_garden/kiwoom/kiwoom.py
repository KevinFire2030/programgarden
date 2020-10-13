from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import *
from pro_garden.config.errorCode import *
from PyQt5.QtTest import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print("kiwoom() class start. ")

        ####### event loop를 실행하기 위한 변수모음
        self.login_event_loop = QEventLoop()  # 로그인 요청용 이벤트루프
        self.detail_account_info_event_loop = QEventLoop()
        self.calculator_event_loop = QEventLoop()

        #########################################

        ####### 변수모음
        self.account_num = None
        self.use_money = 0
        self.use_money_percent = 0.5
        self.account_stock_dict = {}
        self.not_account_stock_dict = {}
        self.calcul_data = []

        #########################################

        ####### 스크린 번호 모음
        self.screen_my_info = "2000"
        self.screen_calculation_stock = "4000"

        #########################################

        ######### 초기 셋팅 함수들 바로 실행
        self.get_ocx_instance()  # OCX 방식을 파이썬에 사용할 수 있게 변환해 주는 함수
        self.event_slots()  # 키움과 연결하기 위한 시그널 / 슬롯 모음
        self.signal_login_commConnect()  # 로그인 요청 시그널 포함
        self.get_account_info()
        self.detail_account_info()
        self.detail_account_mystock()
        self.not_concluded_account()
        self.calculator_fnc()
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

    def not_concluded_account(self, sPrevNext="0"):
        print("미체결요청 부분")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "실시간미체결요청", "opt10075", sPrevNext, self.screen_my_info)

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

        elif sRQName == "계좌평가잔고내역요청":
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

                cnt = cnt + 1

            print("가지고 있는 종목 %s" % cnt)

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()

        elif sRQName == "실시간미체결요청":
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            print("미체결종목 갯수 %s" % rows)
            if rows > 0:
                for i in range(rows):
                    code = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, i, "종목번호")
                    code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                    order_no = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문번호")
                    order_status = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, i, "주문상태")
                    order_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, i, "주문수량")
                    order_price = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, i, "주문가격")
                    order_gubun = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, i, "주문구분")
                    not_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, i, "미체결수량")
                    ok_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, i, "체결량")

                    code = code.strip()
                    code_nm = code_nm.strip()
                    order_no = int(order_no.strip())
                    order_status = order_status.strip()
                    order_quantity = int(order_quantity.strip())
                    order_price = int(order_price.strip())
                    order_gubun = order_gubun.strip().lstrip('+').lstrip('-')
                    not_quantity = int(not_quantity.strip())
                    ok_quantity = int(ok_quantity.strip())

                    if order_no in self.not_account_stock_dict:
                        pass
                    else:
                        self.not_account_stock_dict[order_no] = {}

                    self.not_account_stock_dict[order_no].update({"종목코드": code})
                    self.not_account_stock_dict[order_no].update({'종목명': code_nm})
                    self.not_account_stock_dict[order_no].update({'주문번호': order_no})
                    self.not_account_stock_dict[order_no].update({'주문상태': order_status})
                    self.not_account_stock_dict[order_no].update({'주문수량': order_quantity})
                    self.not_account_stock_dict[order_no].update({'주문가격': order_price})
                    self.not_account_stock_dict[order_no].update({'주문구분': order_gubun})
                    self.not_account_stock_dict[order_no].update({'미체결수량': not_quantity})
                    self.not_account_stock_dict[order_no].update({'체결량': ok_quantity})

                    print("미체결 종목 : %s" % self.not_account_stock_dict[order_no] )
            else:
                print("미체결 종목이 없습니다.")

            self.detail_account_info_event_loop.exit()

        elif sRQName == "주식일봉차트조회요청":
            code = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, 0, "종목코드")
            code = code.strip()
            print("%s 일봉데이터 요청" % code)

            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            print(cnt)

            # 한번에 가져오는 함수
            # data = self.dynamicCall("GetCommDataEx(QString, QString)", sTrCode, sRQName)

            # 한번 조회하면 600일치까지 일봉데이타를 받을 수 있다.
            for i in range(cnt):
                data = []   # 저장할 리스트

                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, i, "현재가")
                value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래량")
                trading_value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래대금")
                date = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "일자")
                start_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "시가")
                high_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "고가")
                low_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "저가")

                data.append("")
                data.append(current_price.strip())
                data.append(value.strip())
                data.append(trading_value.strip())
                data.append(date.strip())
                data.append(start_price.strip())
                data.append(high_price.strip())
                data.append(low_price.trip())
                data.append("")

                self.calcul_data.append(data.copy())

            print(self.calcul_data)

            if sPrevNext == "2":
                self.day_kiwoom_db(code=code, sPrevNext=sPrevNext)
            else:
                print("총 일수 %s" % len(self.calcul_data))

                pass_success = True

                # 120일 이평선을 그릴만큼의 데이터가 있는지 확인
                if self.calcul_data is None or len(self.calcul_data) < 120:
                    pass_success = False
                else:
                    # 데이터가 120일 이상 있으면,
                    total_price = 0
                    for value in self.calcul_data[:120]:
                        total_price += value[1]

                    moving_average_price = total_price / 120

                    # 1. 오늘자 주가가 120 이평선에 걸쳐있는지 확인
                    bottom_stock_price = False
                    check_price = None

                    if int(self.calcul_data[0][7]) <= moving_average_price and\
                        int(self.calcul_data[0][6] >= moving_average_price):
                        print("오늘 주가가 120 이평선에 걸쳐있는지 확인")
                        bottom_stock_price = True
                        check_price = int(self.calcul_data[0][6])

                    # 2. 과거 일봉들이 120일 이평선보다 밑에 있는지 확인
                    prev_low_price = 0  # 과거 일봉 주가
                    if bottom_stock_price is True:
                        moving_average_price_prev = 0
                        price_top_moving = False    # 주가가 이평선보다 위에 위치하는가?

                        idx = 1
                        while True:
                            if len(self.calcul_data[idx:]) < 120:   # 120일자가 있는지 계속 확인
                                print("120일치가 없음")
                                break

                            total_price = 0
                            for value in self.calcul_data[idx:120+idx]:
                                total_price += int(value[1])
                            moving_average_price_prev = total_price / 120

                            if moving_average_price_prev <= int(self.calcul_data[idx][6]) and idx <= 20:
                                print("20일 동안 주가가 이평선과 같거나 위에 있으면 조건 탈락")
                                price_top_moving = False
                                break
                            elif int(self.calcul_data[idx][7]) > moving_average_price_prev and idx > 20:
                                print("120일 이평선 위에 있는 일봉 확인")
                                price_top_moving = True
                                prev_low_price = int(self.calcul_data[idx][7])  # 이평선위의 일봉 저가 저장
                                break

                            idx += 1

                            # 해당 부분 이평선이 가장 최근 일자의 이평선 가격보다 낮은지 확인
                            if price_top_moving is True:
                                if moving_average_price > moving_average_price_prev and check_price > prev_low_price:
                                    print("포착된 이평선의 가격이 오늘자(최근일자) 이평선 가격보다 낮은 것 확임됨")
                                    print("포착된 부분의 일봉 저가가 오늘자 일봉의 고가보다 낮은지 확임됨")
                                    pass_success = True

                    if pass_success is True:
                        pass










                self.calculator_event_loop.exit()

    def day_kiwoom_db(self, code=None, date=None, sPrevNext="0"):
        QTest.qWait(3600)

        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")
        if date is not None:
            self.dynamicCall("SetInputValue(QString, QString)", "기준일자", date)

        self.dynamicCall("CommRqData(QString, QString, int, QString)",
                         "주식일봉차트조회요청", "opt10081", sPrevNext, self.screen_calculation_stock)

        # event loop로 calculator_fnc의 for loop를 멈춰야 함
        self.calculator_event_loop.exec_()

    def get_code_list_by_market(self, market_code):
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market_code)
        code_list = code_list.split(";")[:-1]

        return code_list    # 리스트로 종목코드 저장

    def calculator_fnc(self):
        code_list = self.get_code_list_by_market("10")
        print("코스닥 갯수 %s" % len(code_list))

        for idx, code in enumerate(code_list):
            self.dynamicCall("DisconnectRealData(QString)", self.screen_calculation_stock)  # 스크린번호 해제
            print("%s / %s : KOSDAQ Stock Code : %s is updating..." % (idx+1, len(code_list), code))

            self.day_kiwoom_db(code=code)




