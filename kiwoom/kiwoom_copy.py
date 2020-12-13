import os

from config.kiwoomType import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *
from PyQt5.QtTest import  *
import sys

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print("Kiwoom 클래스")

        self.realType = RealType()

        ######## event loop 모음 #######
        self.login_event_loop = None
        # self.detail_account_info_event_loop = None
        self.detail_account_info_event_loop = QEventLoop()# 다음페이지 있는 경우 스레드 오류 방지
        self.calculator_event_loop = QEventLoop()

        ######## 스크린번호 모음 #######
        self.screen_my_info = "2000"
        self.screen_calculation_stock = "4000"
        self.screen_real_stock = "5000" # 종목별로 할당할 스크린 번호
        self.screen_meme_stock = "6000" # 종목별할당할 주문용 스크린 번호
        self.screen_start_stop_real = "1000"

        ######## 변수 모음 #######
        self.account_num = None

        ######## 종목 분석용 #######
        self.calcul_data = []

        ######## 계좌 관련 변수 ########
        self.use_money = 0
        self.use_money_percent = 0.5

        ########## 변수 모음 ##########
        self.account_stock_dict = {}
        self.portfolio_stock_dict = {}
        self.not_account_stock_dict = {}
        self.jango_dict = {}

        ########## 실행함수 모음 ##########
        self.get_ocx_instance()
        self.event_slots()
        self.real_event_slots()

        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info() # 예수금 가져오기
        self.detail_account_mystock() # 계좌평가잔고 요청
        self.not_concluded_account() # 실시간 미체결 현황 요청
        # self.calculator_fnc() # 종목 분석용, 임시용으로 실행

        self.read_code() # 저장된 종목들 불러오기
        self.screen_number_setting() # 스크린 번호를 할당

        # 실시간 장운영 구분 불러오기
        self.dynamicCall("SetRealReg(QString, QString, QString, QString)", self.screen_start_stop_real, '', self.realType.REALTYPE['장시작시간']['장운영구분'], "0")

        for code in self.portfolio_stock_dict.keys():
            screen_num = self.portfolio_stock_dict[code]['스크린번호']
            fids = self.realType.REALTYPE['주식체결']['체결시간'] # 주식체결에 있는 fid다 가져옴
            self.dynamicCall("SetRealReg(QString, QString, QString, QString)", screen_num, code, fids, "1")
            print("실시간 감시할 종목 코드: %s, 스크린번호: %s fid번호: %s" % (code, screen_num, fids))



    # 윈도우 레지스트리에 등록되어있는 OpenAPI를 사용할 수 있게 등록
    # (PyQt5.QAxContainer 임포트하고 QAxWidget를 상속해서 init 하고 사용)
    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1") # 응용프로그램이 api를 제어할수 있게 해준다.

    def event_slots(self):
        # 원하는 슬롯에 이벤트를 연결(?)해 준다. 이벤트 매핑
        # KOA 개발가이드 보면 OnEventConnect( long nErrCode 를  결과값으로 슬롯으로 넘겨준다. )
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)
        self.OnReceiveMsg.connect(self.msg_slot)

    def real_event_slots(self):
        self.OnReceiveRealData.connect(self.realdata_slot)
        self.OnReceiveChejanData.connect(self.chejan_slot)

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

    # def testWrite(self):
    #     f = open("files/condition_stock.txt", "a", encoding="utf8")
    #     f.write("%s\t%s\t%s\n" % ("016360", "삼성증권", "30000"))
    #     f.close()

    def read_code(self):
        if os.path.exists("files/condition_stock.txt"):
            f = open("files/condition_stock.txt", "r", encoding="utf8")

            lines = f.readlines()
            for line in lines:
                if line != "":
                    ls = line.split("\t")

                    stock_code = ls[0]
                    stock_name = ls[1]
                    stock_price = int(ls[2].split("\n")[0])
                    stock_price = abs(stock_price)

                    self.portfolio_stock_dict.update({stock_code: {"종목명": stock_name, "현재가": stock_price}})

            f.close()

            print("read_code() %s: " % self.portfolio_stock_dict)

    def screen_number_setting(self):

        screen_overwrite = []

        # 계좌평가잔고내역에 있는 종목들
        for code in self.account_stock_dict.keys():
            if code not in screen_overwrite:
                screen_overwrite.append(code)
        # 미체결에 있는 종목들
        for order_no in self.not_account_stock_dict.keys():
            code = self.not_account_stock_dict[order_no]['종목코드']

            if code not in screen_overwrite:
                screen_overwrite.append(code)

        # 포트폴리오에 담긴 종목들
        for code in self.portfolio_stock_dict.keys():
            if code not in screen_overwrite:
                screen_overwrite.append(code)

        # 스크린번호 할당
        cnt = 0
        for code in screen_overwrite:
            temp_screen = int(self.screen_real_stock)
            meme_screen = int(self.screen_meme_stock)

            if (cnt % 50) == 0:
                temp_screen += 1 # 50개씩 넣어서 스크린번호 5000에 종목 50개 할당후 5001에 50개할당
                self.screen_real_stock = str(temp_screen)

            if (cnt % 50) == 0:
                meme_screen += 1
                self.screen_meme_stock = str(meme_screen)

            if code in self.portfolio_stock_dict.keys():
                self.portfolio_stock_dict[code].update({"스크린번호": str(self.screen_real_stock)})
                self.portfolio_stock_dict[code].update({"주문용스크린번호": str(self.screen_meme_stock)})

            elif code not in self.portfolio_stock_dict:
                self.portfolio_stock_dict.update({code: {"스크린번호": str(self.screen_real_stock), "주문용스크린번호": str(self.screen_meme_stock)}})

            cnt += 1
        print("screen_number_setting() %s: " %self.portfolio_stock_dict)

    def realdata_slot(self, sCode, sRealType, sRealData):
        if sRealType == "장시작시간":
            fid = self.realType.REALTYPE[sRealType]['장운영구분']
            value = self.dynamicCall("GetCommRealData(QString, int)", sCode, fid)

            if value == '0':
                print("장 시작전")
            elif value == "3":
                print("장 시작")
            elif value == "2":
                print("장 종료, 동시호가로 넘어감")
            elif value == "4":
                print("3시30분 장 종료")

                for code in self.portfolio_stock_dict.keys():
                    self.dynamicCall("SetRealRemove(QString, QString)", self.portfolio_stock_dict[code]['스크린번호'], code)

                QTest.qWait(5000)

                self.file_delete()
                self.calculator_fnc()

                sys.exit()

        elif sRealType == "주식체결":
            print("실시간 주식체결 코드(틱봉생성) %s : " % sCode) # 틱봉이 생길때마다 코드 찍힘
            a = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['체결시간']) # HHMMSS
            b = abs(int(self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['현재가'])))  # +(-)1000
            c = abs(int(self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['전일대비'])))
            d = float(self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['등락율']))
            e = abs(int(self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['(최우선)매도호가'])))
            f = abs(int(self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['(최우선)매수호가'])))
            g = abs(int(self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['거래량'])))
            h = abs(int(self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['누적거래량'])))
            i = abs(int(self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['고가'])))
            j = abs(int(self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['시가'])))
            k = abs(int(self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['저가'])))

            if sCode not in self.portfolio_stock_dict:
                self.portfolio_stock_dict.update({sCode: {}})

            self.portfolio_stock_dict[sCode].update({'체결시간': a})
            self.portfolio_stock_dict[sCode].update({'현재가': b})
            self.portfolio_stock_dict[sCode].update({'전일대비': c})
            self.portfolio_stock_dict[sCode].update({'등락율': d})
            self.portfolio_stock_dict[sCode].update({'(최우선)매도호가': e})
            self.portfolio_stock_dict[sCode].update({'(최우선)매수호가': f})
            self.portfolio_stock_dict[sCode].update({'거래량': g})
            self.portfolio_stock_dict[sCode].update({'누적거래량': h})
            self.portfolio_stock_dict[sCode].update({'고가': i})
            self.portfolio_stock_dict[sCode].update({'시가': j})
            self.portfolio_stock_dict[sCode].update({'저가': k})

            print(self.portfolio_stock_dict[sCode]) # 장중에 변하는 종목 감시

            # 실시간 조건 검색
            # 계좌잔고평가내역에 있고 오늘 산 잔고에는 없는 경우
            if sCode in self.account_stock_dict.keys() and sCode not in self.jango_dict.keys():
                print("%s %s" % ("신규매도를 한다", sCode))
                asd = self.account_stock_dict[sCode]

                meme_rate = (b - asd['매입가']) / asd['매입가'] * 100

                if asd['매매가능수량'] > 0 and (meme_rate > 5 or meme_rate < -5):
                    order_success = self.dynamicCall("SendOrder(QSting, QSting, QSting, int, QSting, int, int, QSting, QSting)",
                                                     ["신규매도요청", self.portfolio_stock_dict[sCode]['주문용스크린번호'], self.account_num, 2,
                                                     sCode, asd['매매가능수량'], 0, self.realType.SENDTYPE['거래구분']['시장가'], ''])
                    if order_success == 0:
                        print("매도주문 전달 성공")
                        del self.account_stock_dict[sCode]
                    else:
                        print("매도주문 전달 실패")

            # 오늘 산 잔고에 있을 경우
            elif sCode in self.jango_dict.keys():
                print("%s %s" % ("신규매도를 한다2", sCode))

                jd = self.jango_dict[sCode]
                meme_rate = (b - jd['매입단가']) / jd['매입단가'] * 100

                if jd['주문가능수량'] > 0 and (meme_rate > 5 or meme_rate < -5):

                    order_success = self.dynamicCall(
                        "SendOrder(QSting, QSting, QSting, int, QSting, int, int, QSting, QSting)",
                        ["신규매도요청", self.portfolio_stock_dict[sCode]['주문용스크린번호'], self.account_num, 2, sCode,
                         jd['주문가능수량'], 0, self.realType.SENDTYPE['거래구분']['시장가'], '']
                    )

                    if order_success == 0:
                        print("매도주문 전달 성공")
                    else:
                        print("매도주문 전달 실패")

            # 등락율이 2.0이상이고 오늘 산 잔고에 없는 경우
            elif d > 2.0 and sCode not in self.jango_dict:
                print("%s %s" % ("신규매수를 한다", sCode))

                result = (self.use_money * 0.1) / e
                quantity = int(result)

                order_success = self.dynamicCall(
                    "SendOrder(QSting, QSting, QSting, int, QSting, int, int, QSting, QSting)",
                    ['신규매수', self.portfolio_stock_dict[sCode]['주문용스크린번호'], self.account_num,
                     1, sCode, quantity, e, self.realType.REALTYPE['거래구분']['지정가'], '']
                )

                if order_success == 0:
                    print("매수주문 전달 성공")
                else:
                    print("매수주문 전달 실패")

            # list를 씌워서 같은 주소를 바라보지 않게 한다.(복사본개념) / 계산하는동안에 체결될수있으므로 복사한다.
            not_meme_list = list(self.not_account_stock_dict)
            for order_num in not_meme_list:
                code = self.not_account_stock_dict[order_num]['종목코드']
                meme_price = self.not_account_stock_dict[order_num]['주문가격']
                not_quantity = self.not_account_stock_dict[order_num]['미체결수량']
                order_gubun = self.not_account_stock_dict[order_num]['주문구분']

                if order_gubun == "매수" and not_quantity > 0 and e > meme_price:
                    print("%s %s" % ("매수취소한다.", sCode))
                    order_success = self.dynamicCall(
                        "SendOrder(QSting, QSting, QSting, int, QSting, int, int, QSting, QSting)",
                        ["매수취소", self.portfolio_stock_dict[sCode]['주문용스크린번호'], self.account_num, 3,
                         code, 0, 0, self.realType.SENDTYPE['거래구분']['지정가'], order_num]
                    )

                    if order_success == 0:
                        print("매수취소 전달 성공")
                    else:
                        print("매수취소 전다 실패")

                elif not_quantity == 0:
                    del self.not_account_stock_dict[order_num]

    # sendOrder 된 결과가 여기로 옴
    def chejan_slot(self, sGubun, nItemCnt, sFIdList):

        if int(sGubun) == 0:
            print("주문체결 슬롯 전달")
            realTypeOrder = self.realType.REALTYPE['주문체결']
            account_num = self.dynamicCall("GetChejanData(int)", realTypeOrder['계좌번호'])
            sCode = self.dynamicCall("GetChejanData(int)", realTypeOrder['종목코드'])
            stock_name = self.dynamicCall("GetChejanData(int)", realTypeOrder['종목명']).strip()
            origin_order_number = self.dynamicCall("GetChejanData(int)", realTypeOrder['원주문번호'])
            order_number = self.dynamicCall("GetChejanData(int)", realTypeOrder['주문번호'])
            order_status = self.dynamicCall("GetChejanData(int)", realTypeOrder['주문상태'])
            order_quan = int(self.dynamicCall("GetChejanData(int)", realTypeOrder['주문수량']))
            order_price = int(self.dynamicCall("GetChejanData(int)", realTypeOrder['주문가격']))
            not_chegual_quan = int(self.dynamicCall("GetChejanData(int)", realTypeOrder['미체결수량']))
            order_gubun = self.dynamicCall("GetChejanData(int)", realTypeOrder['주문구분']).strip().lstrip('+').lstrip('-')
            chegual_time_str = self.dynamicCall("GetChejanData(int)", realTypeOrder['주문/체결시간'])
            chegual_price = self.dynamicCall("GetChejanData(int)", realTypeOrder['체결가'])
            if chegual_price == '':
                chegual_price = 0
            else:
                chegual_price = int(chegual_price)

            chegual_quantity = self.dynamicCall("GetChejanData(int)", realTypeOrder['체결량'])
            if chegual_quantity == '':
                chegual_quantity = 0
            else:
                chegual_quantity = int(chegual_quantity)

            current_price = abs(int(self.dynamicCall("GetChejanData(int)", realTypeOrder['현재가'])))
            first_sell_price = abs(int(self.dynamicCall("GetChejanData(int)", realTypeOrder['(최우선)매도호가'])))
            first_buy_price = abs(int(self.dynamicCall("GetChejanData(int)", realTypeOrder['(최우선)매수호가'])))

            # 새로 들어온 주문이면 주문번호 할당
            if order_number not in self.not_account_stock_dict.keys():
                self.not_account_stock_dict.update({order_number: {}})

            self.not_account_stock_dict[order_number].update({"종목코드": sCode})
            self.not_account_stock_dict[order_number].update({"주문번호": order_number})
            self.not_account_stock_dict[order_number].update({"종목명": stock_name})
            self.not_account_stock_dict[order_number].update({"주문상태": order_status})
            self.not_account_stock_dict[order_number].update({"주문수량": order_quan})
            self.not_account_stock_dict[order_number].update({"주문가격": order_price})
            self.not_account_stock_dict[order_number].update({"미체결수량": not_chegual_quan})
            self.not_account_stock_dict[order_number].update({"원주문번호": origin_order_number})
            self.not_account_stock_dict[order_number].update({"주문구분": order_quan})
            self.not_account_stock_dict[order_number].update({"주문/체결시간": chegual_time_str})
            self.not_account_stock_dict[order_number].update({"체결가": chegual_price})
            self.not_account_stock_dict[order_number].update({"체결량": chegual_quantity})
            self.not_account_stock_dict[order_number].update({"현재가": current_price})
            self.not_account_stock_dict[order_number].update({"(최우선)매도호가": first_sell_price})
            self.not_account_stock_dict[order_number].update({"(최우선)매수호가": first_buy_price})

            print(self.not_account_stock_dict)

        elif int(sGubun) == 1:
            print("잔고 슬롯 전달")
            realTypeJango = self.realType.REALTYPE['잔고']

            account_num = self.dynamicCall("GetChejanData(int)", realTypeJango['계좌번호'])
            sCode = self.dynamicCall("GetChejanData(int)", realTypeJango['종목코드'])[1:]
            stock_name = self.dynamicCall("GetChejanData(int)", realTypeJango['종목명']).strip()
            current_price = abs(int(self.dynamicCall("GetChejanData(int)", realTypeJango['현재가'])))
            stock_quan = int(self.dynamicCall("GetChejanData(int)", realTypeJango['보유수량']))
            like_quan = int(self.dynamicCall("GetChejanData(int)", realTypeJango['주문가능수량']))
            buy_price = abs(int(self.dynamicCall("GetChejanData(int)", realTypeJango['매입단가'])))
            total_buy_price = int(self.dynamicCall("GetChejanData(int)", realTypeJango['총매입가']))
            meme_gubun = self.dynamicCall("GetChejanData(int)", realTypeJango['매도매수구분'])
            meme_gubun = self.realType.REALTYPE['매도수구분'][meme_gubun]
            first_sell_price = abs(int(self.dynamicCall("GetChejanData(int)", realTypeJango['(최우선)매도호가'])))
            first_buy_price = abs(int(self.dynamicCall("GetChejanData(int)", realTypeJango['(최우선)매수호가'])))

            if sCode not in self.jango_dict.keys():
                self.jango_dict.update({sCode:{}})

            self.jango_dict[sCode].update({"현재가": current_price})
            self.jango_dict[sCode].update({"종목코드": sCode})
            self.jango_dict[sCode].update({"종목명": stock_name})
            self.jango_dict[sCode].update({"보유수량": stock_quan})
            self.jango_dict[sCode].update({"주문가능수량": like_quan})
            self.jango_dict[sCode].update({"매입단가": buy_price})
            self.jango_dict[sCode].update({"총매입가": total_buy_price})
            self.jango_dict[sCode].update({"매도매수구분": meme_gubun})
            self.jango_dict[sCode].update({"(최우선)매도호가": first_sell_price})
            self.jango_dict[sCode].update({"(최우선)매수호가": first_buy_price})

            if stock_quan == 0:
                del self.jango_dict[sCode]
                self.dynamicCall("setRealRemove(QString, QString)", self.portfolio_stock_dict[sCode]['스크린번호'], sCode)

    # 송수신 메세지 get
    def msg_slot(self, sScrNo, sRQName, sTrCode, msg):
        print("스크린: %s, 요청이름: %s , tr코드: %s --- %s" % (sScrNo, sRQName, sTrCode, msg))


    # 분석 파일 삭제
    def file_delete(self):
        if os.path.isfile("files/condition_stock.txt"):
            os.remove("files/condition_stock.txt")






