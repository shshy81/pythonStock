from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *
from PyQt5.QtTest import  *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print("Kiwoom 클래스")
        ######## event loop 모음 #######
        self.login_event_loop = None
        # self.detail_account_info_event_loop = None
        self.detail_account_info_event_loop = QEventLoop()# 다음페이지 있는 경우 스레드 오류 방지
        self.calculator_event_loop = QEventLoop()

        ######## 스크린번호 모음 #######
        self.screen_my_info = "2000"
        self.screen_calculation_stock = "4000"

        ######## 변수 모음 #######
        self.account_num = None
        self.not_account_stock_dict = {}

        ######## 종목 분석용 #######
        self.calcul_data = []

        ######## 계좌 관련 변수 ########
        self.use_money = 0
        self.use_money_percent = 0.5

        ########## 변수 모음 ##########
        self.account_stock_dict = {}

        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info() # 예수금 가져오기
        self.detail_account_mystock() # 계좌평가잔고 요청
        self.not_concluded_account() # 실시간 미체결 현황 요청
        self.calculator_fnc() # 종목 분석용, 임시용으로 실행

    # 윈도우 레지스트리에 등록되어있는 OpenAPI를 사용할 수 있게 등록
    # (PyQt5.QAxContainer 임포트하고 QAxWidget를 상속해서 init 하고 사용)
    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1") # 응용프로그램이 api를 제어할수 있게 해준다.

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        # 원하는 슬롯에 이벤트를 연결(?)해 준다. 이벤트 매핑
        # KOA 개발가이드 보면 OnEventConnect( long nErrCode 를  결과값으로 슬롯으로 넘겨준다. )

        self.OnReceiveTrData.connect(self.trdata_slot)

    def signal_login_commConnect(self): # login 요청
        self.dynamicCall("CommConnect") # PyQt5 에서 제공. / 네트워크, 서버, 응용프로그램에 데이터를 전송하는 함수

        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def login_slot(self, errCode):
        print(errors(errCode)) # 여기선 왜 self.errors로 안하지?

        self.login_event_loop.exit()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        self.account_num = account_list.split(';')[0]
        print("나의 계좌번호 %s" % self.account_num) #8151481111

    def detail_account_info(self):
        print("예수금 요청")

        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)", "예수금상세현황요청", "opw00001", "0", self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):
        print("계좌평가잔고내역요청")

        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)", "계좌평가잔고내역요청", "opw00018", sPrevNext, self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    def not_concluded_account(self, sPrevNext="0"):
        print("실시간미체결요청")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "전체종목구분", "0")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", "")
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "실시간미체결요청", "opt10075", sPrevNext, self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
        tr요청을 받는 슬롯구역
        :param sScrNo: 스크린번호
        :param sRQName: 요청이름
        :param sTrCode: 요청id, tr코드
        :param sRecordName: 레코드 이름
        :param sPrevNext: 다음 페이지가 있는지
        :return:
        '''

        if sRQName == '예수금상세현황요청':
            deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "예수금")
            print("예수금 %s" % int(deposit))

            self.use_money = int(deposit) * self.use_money_percent
            self.use_money = self.use_money / 4

            posibleDeposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "출금가능금액")
            print("출금가능금액 %s" % int(posibleDeposit))

            self.detail_account_info_event_loop.exit()

        elif sRQName == '계좌평가잔고내역요청':
            total_buy_amount = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총매입금액")
            print("총매입금액 %s" % int(total_buy_amount))

            total_profit_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총수익률(%)")
            print("총수익률(%%) %s" % float(total_profit_rate))

            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            cnt = 0
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "종목번호")
                code = code.strip()[1:]
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "종목명")
                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "보유수량")
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "매입가")
                return_rate = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "수익률(%)")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "현재가")
                total_chegual_price = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "매입금액")
                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "매매가능수량")

                if code in self.account_stock_dict:
                    pass
                else:
                    self.account_stock_dict.update({code: {}})

                account_dict = self.account_stock_dict[code]
                code_nm = code_nm.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                return_rate = float(return_rate.strip())
                current_price = int(current_price.strip())
                total_chegual_price = int(total_chegual_price.strip())
                possible_quantity = int(possible_quantity.strip())

                account_dict.update({"종목명": code_nm})
                account_dict.update({"보유수량": stock_quantity})
                account_dict.update({"매입가": buy_price})
                account_dict.update({"수익률(%)": return_rate})
                account_dict.update({"현재가": current_price})
                account_dict.update({"매입금액": total_chegual_price})
                account_dict.update({"매매가능수량": possible_quantity})

                cnt += 1

            print("보유 종목 갯수 %s" % cnt)
            # print("보유 종목 %s" % self.account_stock_dict)

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext = "2")
            else:
                self.detail_account_info_event_loop.exit()

        elif sRQName == "실시간미체결요청":
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "종목코드")
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "종목명")
                order_no = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "주문번호")
                order_status = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "주문상태")
                order_quantity = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "주문수량")
                order_price = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "주문가격")
                order_gubun = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "주문구분")
                not_quantity = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "미체결수량")
                ok_quantity = self.dynamicCall("GetCommData(QString, QString, int, String)", sTrCode, sRQName, i, "체결량")

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

                not_account_stock_dict = self.not_account_stock_dict[order_no]

                not_account_stock_dict.update({'종목코드': code})
                not_account_stock_dict.update({'종목명': code_nm})
                not_account_stock_dict.update({'주문번호': order_no})
                not_account_stock_dict.update({'주문상태': order_status})
                not_account_stock_dict.update({'주문수량': order_quantity})
                not_account_stock_dict.update({'주문가격': order_price})
                not_account_stock_dict.update({'주문구분': order_gubun})
                not_account_stock_dict.update({'미체결수량': not_quantity})
                not_account_stock_dict.update({'체결량': ok_quantity})

                print("미체결 종목 : %s " % not_account_stock_dict)

            self.detail_account_info_event_loop.exit()
        elif sRQName == "주식일봉차트조회요청":
            code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "종목코드")
            code = code.strip()
            print("%s 일봉데이터 요청" % code)

            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            print("데이터 일수 %s" % cnt)

            # data = self.dynamicCall("GetCommDataEx(QString, QString)", sTrCode, sRQName)
            # 결과값 => [['', '현재가', '거래량', '거래대금', '날짜', '시가', '고가', '저가', ''], ['', '현재가', '거래량', '거래대금', '날짜', '시가', '고가', '저가', '']]

            for i in range(cnt):
                data = []

                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
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
                data.append(low_price.strip())
                data.append("")

                self.calcul_data.append(data.copy())

            if sPrevNext == "2":
                self.day_kiwoom_db(code=code, sPrevNext=sPrevNext)
            else:
                print("총 일수 %s" % len(self.calcul_data))

                pass_success = False
                if self.calcul_data == None or len(self.calcul_data) < 120:
                    pass_success = False
                else:
                    total_price = 0
                    for value in self.calcul_data[:120]:
                        total_price += int(value[1]) # 현재가(종가)
                    moving_average_price = total_price / 120
                    # 오늘자 주가가 120일에 걸쳐있는지 확인
                    bottom_stock_price = False
                    check_price = None
                    if int(self.calcul_data[0][7]) <= moving_average_price and moving_average_price <= int(self.calcul_data[0][6]):
                        print("오늘 주가 120이평선에 걸쳐있는것 확인")
                        bottom_stock_price = True
                        check_price = int(self.calcul_data[0][6]) # 오늘 고가
                    # 과거 일봉들이 120일 이평선 아래 있는지 확인, 일봉이 120일 이평선보다 위에 있으면 계산 진행

                    prev_price = None # 과거의 일봉 저가
                    if bottom_stock_price == True:
                        moving_average_price_prev = 0
                        price_top_moving = False

                        idx = 1
                        while True:
                            if len(self.calcul_data[idx:]) < 120: #120일치가 있는지 계속 확인
                                print("120일치가 없음")
                                break
                            total_price = 0
                            for value in self.calcul_data[idx:120+idx]:
                                total_price += int(value[1])
                            moving_average_price_prev = total_price / 120

                            if moving_average_price_prev <= int(self.calcul_data[idx][6]) and idx <= 20:
                                print("20일 동안 주가가 120일 이평선과 같거나 위에 있으면 조건 통과 못함")
                                price_top_moving = False
                                break

                            elif int(self.calcul_data[idx][7]) > moving_average_price_prev and idx > 20:
                                print("120일 이평선 위에 있는 일봉 확인됨")
                                price_top_moving = True
                                prev_price = int(self.calcul_data[idx][7])
                                break

                            idx += 1

                        # 해당 부분 이평선이 가장 최근 일자의 이평선 가격보다 낮은지 확인
                        if price_top_moving == True:
                            if moving_average_price > moving_average_price_prev and check_price > prev_price:
                                print("포착된 이평선의 가격이 오늘자(최근일자) 이평선 가격보다 낮은것 확인됨")
                                print("포착된 부분의 일봉 저가가 오늘자 일봉의 고가보다 낮은것 확인됨")
                                pass_success = True

                if pass_success == True:
                    print("조건부 통과됨")

                    code_nm = self.dynamicCall("GetMasterCodeName(QString)", code)

                    f = open("files/condition_stock.txt", "a", encoding="utf8")
                    f.write("%s\t%s\t%s\n" % (code, code_nm, str(self.calcul_data[0][1])))
                    f.close()

                elif pass_success == False:
                    print("조건부 통과 못함")

                self.calcul_data.clear()
                self.calculator_event_loop.exit()

    def get_code_list_by_market(self, market_code):
        '''
        종목 코드들 반환
        :param market_code:
        :return:
        '''
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market_code)
        code_list = code_list.split(";")[:-1]

        return code_list

    def calculator_fnc(self):
        '''
        종목 분석 실행용 함수
        :return:
        '''
        code_list = self.get_code_list_by_market("10")
        print("코스닥 갯수 %s" % len(code_list))

        for idx, code in enumerate(code_list):
            self.dynamicCall("DisconnectRealData(QString)", self.screen_calculation_stock)

            print("%s / %s : KOSDAQ Stock Code : %s is updating ... " % (idx + 1, len(code_list), code))

            self.day_kiwoom_db(code = code)

    def day_kiwoom_db(self, code=None, date=None, sPrevNext="0"):

        QTest.qWait(3600)

        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")
        if date != None:
            self.dynamicCall("SetInputValue(QString, QString)", "기준일자", date)
        self.dynamicCall("CommRqData(QString, QString, int, QString", "주식일봉차트조회요청", "opt10081", sPrevNext, self.screen_calculation_stock)

        self.calculator_event_loop.exec_()
