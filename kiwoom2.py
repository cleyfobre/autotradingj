from datetime import datetime
from datetime import timedelta
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from app.kiwoom_types import *
import logging
from pprint import pformat

class Kiwoom(QAxWidget):

    def __init__(self):
        super().__init__()

        logging.basicConfig(format="%(asctime)s | %(message)s", level=logging.NOTSET)
        logging.info("키움 자동 매매 어플리케이션 시작\n")

        # CHECK THESE VARIABLES BEFORE RUNNING!!
        # MODE: PROD | QA
        self.MODE = "QA"
        self.USE_RATIO = 0.01
        self.BUYING_AMOUNT = 50000
        self.TIME_TO_START_SEARCHING = "09:01:00"
        self.TIME_TO_STOP_SEARCHING = "15:18:00"
        self.TIME_TOMODE_START_TO_BUY = "09:01:00"
        self.TIME_TO_SUMMARIZE = "10:11:00"
        self.TIME_TO_STOP_TRADING = "15:18:00"
        self.CONDITION_SEARCH_INDEX = 21 # 검색기 목록 인덱스

        # Kiwoom screen number
        self.SCREEN_CONDITION_SEARCH = "0156"

        # API screen numbers
        self.SCREEN_MY_INFO = "2000"
        self.SCREEN_STOCK_CALCULATION = "4000"
        self.SCREEN_REALTIME = "5000"
        self.SCREEN_TRADING = "6000"
        self.SCREEN_REAL_REG = "1000"

        # final variables
        self.PROD = "PROD"
        self.FILE_PATH = "C:/Users/cleyf/PycharmProjects/autotrading_j/logs/"
        self.LOG_PATH = self.FILE_PATH + datetime.now().strftime("%y%m%d") + "log.csv"
        self.CANDIDATES_PATH = self.FILE_PATH + datetime.now().strftime("%y%m%d") + "candidates.txt"
        self.FAILED_PATH = self.FILE_PATH + datetime.now().strftime("%y%m%d") + "failed.txt"
        self.TARGETS_PATH = self.FILE_PATH + datetime.now().strftime("%y%m%d") + "targets.txt"
        self.SECONDS_1 = 1
        self.SECONDS_2 = 2
        self.SECONDS_5 = 5
        self.SECONDS_10 = 10
        self.SECONDS_60 = 60
        self.SECONDS_1800 = 1800
        # 30min = 1800sec
        self.MINUTES_1 = 60
        self.MINUTES_2 = 120
        self.MINUTES_3 = 180
        self.MINUTES_4 = 240
        self.MINUTES_5 = 300
        self.MINUTES_15 = 900
        self.ETF_ALIAS = ["KODEX ", "TIGER ", "KOSEF ", "KBSTAR "]

        # variables
        self.is_summarized = True
        self.use_amount = 0
        self.account_num = None
        self.holdings_dict = {}
        self.unfulfilled_dict = {}
        self.calculated_data = []
        self.candidates_dict = {}
        self.realType = RealType()
        self.targets_dict = {}
        self.on_trading_list = []
        self.failed_dict = {}

        # event loops
        self.login_event_loop = QEventLoop()
        self.kiwoon_event_loop = QEventLoop()
        self.calculator_event_loop = QEventLoop()
        self.condition_search_event_loop = QEventLoop()

        # callbacks
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")
        self.OnEventConnect.connect(self._login_handler)
        self.OnReceiveConditionVer.connect(self._condition_ver_handler)
        self.OnReceiveTrCondition.connect(self._condition_search_handler)
        self.OnReceiveRealCondition.connect(self._realtime_condition_search_handler)
        self.OnReceiveTrData.connect(self._trdata_handler)
        self.OnReceiveRealData.connect(self._real_data_handler)  # 실시간 이벤트 연결
        self.OnReceiveChejanData.connect(self._chejan_handler)  # 종목 주문체결 관련한 이벤트

        # trading
        self.today_log()
        self.login()
        self.my_account_info()
        self.my_holdings()
        self.my_unfulfilled()

        self.condition_search()
        self.screen_numbers()
        self.candidates_registration()

    def login(self):
        self.dynamicCall("commConnect()")
        self.login_event_loop.exec_()

    def _login_handler(self, error_code):
        # logging.info(errors(errCode))
        if error_code == 0:
            # old: def account_info(self)
            account_index = 1
            if self.MODE == self.PROD:
                account_index = 0
                logging.info("*** PRODUCTION TRADING\n")
            else:
                logging.info("* QA TRADING\n")
            self.account_num = self.dynamicCall("GetLoginInfo(QString)", "ACCNO").split(";")[account_index]
            logging.info("계좌 번호: %s", self.account_num)
        else:
            logging.info("\n로그인 실패\n")

        self.login_event_loop.exit()

    def my_account_info(self):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", 0000)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", 00)
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        self.dynamicCall("CommRqData(QString, QString, QString, QString)",
                         "예수금상세현황요청", "opw00001", "0", self.SCREEN_MY_INFO)

        self.kiwoon_event_loop.exec_()

    def my_holdings(self, prev_next="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", 0000)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", 00)
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        self.dynamicCall("CommRqData(QString, QString, QString, QString)",
                         "계좌평가잔고내역요청", "opw00018", prev_next, self.SCREEN_MY_INFO)

        self.kiwoon_event_loop.exec_()

    def my_unfulfilled(self, prev_next="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("CommRqData(QString, QString, int, QString)",
                         "실시간미체결요청", "opt10075", prev_next, self.SCREEN_MY_INFO)

        self.kiwoon_event_loop.exec_()

    def _trdata_handler(self, scr_no, rq_name, tr_code, record_name, prev_next):
        """
        Slot to receive tr request
        :param scr_no: Screen number
        :param rq_name: Request name
        :param tr_code: tr code
        :param record_name: no use
        :param prev_next: if there is next page
        :return:
        """

        if rq_name == "예수금상세현황요청":
            balance = int(
                self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, 0, "예수금"))
            logging.info("예수금: %s", balance)
            self.use_amount = balance * self.USE_RATIO

            logging.info("출금 가능 금액: %s", int(
                self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, 0, "출금가능금액")))

            self.kiwoon_event_loop.exit()

        elif rq_name == "계좌평가잔고내역요청":
            logging.info("총 매입 금액: %s", int(
                self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, 0, "총매입금액")))

            logging.info("총 수익률: %s", float(
                self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, 0, "총수익률(%)")))

            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", tr_code, rq_name)
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                        tr_code, rq_name, i, "종목번호").strip()[1:]
                code_name = self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                             tr_code, rq_name, i, "종목명").strip()
                stock_qtt = int(self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                 tr_code, rq_name, i, "보유수량").strip())
                buy_price = int(self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                 tr_code, rq_name, i, "매입가").strip())
                revenue = float(self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                 tr_code, rq_name, i, "수익률(%)").strip())
                current_price = int(self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                     tr_code, rq_name, i, "현재가").strip())
                buy_amount = int(self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                  tr_code, rq_name, i, "매입금액").strip())
                possible_qtt = int(self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                    tr_code, rq_name, i, "매매가능수량").strip())

                if code in self.holdings_dict:
                    pass
                else:
                    self.holdings_dict.update({code: {}})

                self.holdings_dict[code].update({"종목명": code_name})
                self.holdings_dict[code].update({"보유수량": stock_qtt})
                self.holdings_dict[code].update({"매입가": buy_price})
                self.holdings_dict[code].update({"수익률(%)": revenue})
                self.holdings_dict[code].update({"현재가": current_price})
                self.holdings_dict[code].update({"매입금액": buy_amount})
                self.holdings_dict[code].update({"매매가능수량": possible_qtt})

            logging.info("보유 종목: \n%s", pformat(self.holdings_dict))

            if prev_next == "2":
                self.my_holdings(prev_next)
            else:
                self.kiwoon_event_loop.exit()

        elif rq_name == "실시간미체결요청":
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", tr_code, rq_name)
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                        tr_code, rq_name, i, "종목코드").strip()
                code_name = self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                             tr_code, rq_name, i, "종목명").strip()
                order_no = int(self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                tr_code, rq_name, i, "주문번호").strip())
                # 접수,확인,체결
                order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                tr_code, rq_name, i, "주문상태").strip()
                order_qtt = int(self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                 tr_code, rq_name, i, "주문수량").strip())
                order_price = int(self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                   tr_code, rq_name, i, "주문가격").strip())
                # -매도, +매수, -매도정정, +매수정정
                order_cat = self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                               tr_code, rq_name, i, "주문구분").strip().lstrip('+').lstrip('-')
                unconcluded_qtt = int(self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                       tr_code, rq_name, i, "미체결수량").strip())
                concluded_qtt = int(self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                     tr_code, rq_name, i, "체결량").strip())

                if order_no in self.unfulfilled_dict:
                    pass
                else:
                    self.unfulfilled_dict[order_no] = {}

                self.unfulfilled_dict[order_no].update({'종목코드': code})
                self.unfulfilled_dict[order_no].update({'종목명': code_name})
                self.unfulfilled_dict[order_no].update({'주문번호': order_no})
                self.unfulfilled_dict[order_no].update({'주문상태': order_status})
                self.unfulfilled_dict[order_no].update({'주문수량': order_qtt})
                self.unfulfilled_dict[order_no].update({'주문가격': order_price})
                self.unfulfilled_dict[order_no].update({'주문구분': order_cat})
                self.unfulfilled_dict[order_no].update({'미체결수량': unconcluded_qtt})
                self.unfulfilled_dict[order_no].update({'체결량': concluded_qtt})

            logging.info("미체결 종목: \n%s", pformat(self.unfulfilled_dict))

            self.kiwoon_event_loop.exit()

    def candidates_registration(self):
        self.dynamicCall("SetRealReg(QString, QString, QString, QString)", self.SCREEN_REAL_REG, '',
                         self.realType.REAL_TYPE['장시작시간']['장운영구분'], "0")
        logging.info("매매할 종목: \n%s", pformat(self.candidates_dict))
        for code in self.candidates_dict.keys():
            self.dynamicCall("SetRealReg(QString, QString, QString, QString)",
                             self.candidates_dict[code]['화면번호'],
                             code, self.realType.REAL_TYPE['주식체결']['체결시간'], "1")

    def _realtime_condition_search_handler(self, code, event, cond_name, cond_idx):
        """
        실시간 종목 조건검색 요청시 발생되는 이벤트

        :param code: string - 종목코드
        :param event: string - 이벤트종류("I": 종목편입, "D": 종목이탈)
        :param cond_name: string - 조건식 이름
        :param cond_idx: string - 조건식 인덱스(여기서만 인덱스가 string 타입으로 전달됨)
        """

        now = datetime.now().strftime("%H:%M:%S")
        if now > self.TIME_TO_STOP_SEARCHING:
            self.dynamicCall("SendConditionStop(QString, QString, int)", self.SCREEN_CONDITION_SEARCH, cond_name,
                             cond_idx)

        if now > self.TIME_TO_START_SEARCHING and event == "I":
            _name = self.dynamicCall("GetMasterCodeName(QString)", [code])
            temp_screen = int(self.SCREEN_REALTIME)
            meme_screen = int(self.SCREEN_TRADING)

            if (len(self.candidates_dict) % 50) == 0:
                temp_screen += 1
                self.SCREEN_REALTIME = str(temp_screen)

            if (len(self.candidates_dict) % 50) == 0:
                meme_screen += 1
                self.SCREEN_TRADING = str(meme_screen)

            if code in self.candidates_dict.keys():
                if "종목명" not in self.candidates_dict[code]:
                    self.candidates_dict[code].update({"종목명": _name})
                if "화면번호" not in self.candidates_dict[code]:
                    self.candidates_dict[code].update({"화면번호": str(self.SCREEN_REALTIME)})
                if "주문용화면번호" not in self.candidates_dict[code]:
                    self.candidates_dict[code].update({"주문용화면번호": str(self.SCREEN_TRADING)})
                if "captured_time" not in self.candidates_dict[code]:
                    self.candidates_dict[code].update({"captured_time": datetime.now().strftime("%H:%M:%S")})
            elif code not in self.candidates_dict.keys():
                self.candidates_dict.update({
                    code: {
                        "종목명": _name, "화면번호": str(self.SCREEN_REALTIME),
                        "주문용화면번호": str(self.SCREEN_TRADING), "captured_time": datetime.now().strftime("%H:%M:%S")
                    }})

                self.dynamicCall("SetRealReg(QString, QString, QString, QString)",
                                 self.candidates_dict[code]['화면번호'], code,
                                 self.realType.REAL_TYPE['주식체결']['체결시간'], "1")

    def _condition_search_handler(self, scr_no, code_list, cond_name, index, next):
        cls = code_list.split(";")
        for code in cls:
            now = datetime.now().strftime("%H:%M:%S")
            if now > self.TIME_TO_START_SEARCHING and code != "":
                self.candidates_dict.update({code: {}})

        self.condition_search_event_loop.exit()

    def _condition_ver_handler(self):
        cond_list = {"list": [], "name": []}
        temp_cond_list = self.dynamicCall("GetConditionNameList()").split(";")
        logging.info("검색식 리스트: %s", temp_cond_list)

        for data in temp_cond_list:
            try:
                a = data.split("^")
                cond_list["list"].append(str(a[0]))
                cond_list["name"].append(str(a[1]))
            except IndexError:
                pass

        cond_name = cond_list["name"][self.CONDITION_SEARCH_INDEX]
        cond_idx = cond_list["list"][self.CONDITION_SEARCH_INDEX]

        cond_search_success = self.dynamicCall("SendCondition(QString, QString, int, int)",
                                               self.SCREEN_CONDITION_SEARCH, cond_name, cond_idx, 1)
        if cond_search_success != 1:
            logging.info("조건 검색 실패")

    def condition_search(self):
        self.dynamicCall("GetConditionLoad()")
        self.condition_search_event_loop.exec_()

    def screen_numbers(self):
        screen_overwrite = []

        # 계좌평가잔고내역에 있는 종목들
        for code in self.holdings_dict.keys():
            if code not in screen_overwrite:
                screen_overwrite.append(code)

        # 미체결에 있는 종목들
        for order_number in self.unfulfilled_dict.keys():
            code = self.unfulfilled_dict[order_number]['종목코드']

            if code not in screen_overwrite:
                screen_overwrite.append(code)

        # 포트폴리오에 있는 종목들
        for code in self.candidates_dict.keys():
            if code not in screen_overwrite:
                screen_overwrite.append(code)

        # 스크린 번호 할당
        cnt = 0
        for code in screen_overwrite:
            temp_screen = int(self.SCREEN_REALTIME)
            meme_screen = int(self.SCREEN_TRADING)

            if (cnt % 50) == 0:
                temp_screen += 1
                self.SCREEN_REALTIME = str(temp_screen)

            if (cnt % 50) == 0:
                meme_screen += 1
                self.SCREEN_TRADING = str(meme_screen)

            _name = self.dynamicCall("GetMasterCodeName(QString)", [code])
            if code in self.candidates_dict.keys():
                if "종목명" not in self.candidates_dict[code]:
                    self.candidates_dict[code].update({"종목명": _name})
                if "화면번호" not in self.candidates_dict[code]:
                    self.candidates_dict[code].update({"화면번호": str(self.SCREEN_REALTIME)})
                if "주문용화면번호" not in self.candidates_dict[code]:
                    self.candidates_dict[code].update({"주문용화면번호": str(self.SCREEN_TRADING)})
                if "captured_time" not in self.candidates_dict[code]:
                    self.candidates_dict[code].update({"captured_time": datetime.now().strftime("%H:%M:%S")})
            elif code not in self.candidates_dict.keys():
                self.candidates_dict.update({
                    code: {
                        "종목명": _name, "화면번호": str(self.SCREEN_REALTIME),
                        "주문용화면번호": str(self.SCREEN_TRADING), "captured_time": datetime.now().strftime("%H:%M:%S")
                    }})

            cnt += 1

    def _real_data_handler(self, s_code, real_type, real_data):
        if real_type == "장시작시간":
            fid = self.realType.REAL_TYPE[real_type]['장운영구분']
            value = self.dynamicCall("GetCommRealData(QString, int)", s_code, fid)

            if value == '0':
                logging.info("장 시작 전")

            elif value == '3':
                logging.info("장 시작")

            elif value == "2":
                logging.info("장 종료, 동시호가로 넘어감")

            elif value == "4":
                logging.info("3시30분 장 종료")

                for c in self.candidates_dict.keys():
                    self.dynamicCall("SetRealRemove(QString, QString)", self.candidates_dict[c]['화면번호'], c)

        elif real_type == "주식체결":
            concluded_time = self.dynamicCall("GetCommRealData(QString, int)",
                                 s_code, self.realType.REAL_TYPE[real_type]['체결시간'])
            current_price = abs(int(self.dynamicCall("GetCommRealData(QString, int)",
                                         s_code, self.realType.REAL_TYPE[real_type]['현재가'])))
            range_day2day = abs(int(self.dynamicCall("GetCommRealData(QString, int)",
                                         s_code, self.realType.REAL_TYPE[real_type]['전일대비'])))
            today_range = float(self.dynamicCall("GetCommRealData(QString, int)",
                                       s_code, self.realType.REAL_TYPE[real_type]['등락율']))
            ask_price = abs(int(self.dynamicCall("GetCommRealData(QString, int)",
                                         s_code, self.realType.REAL_TYPE[real_type]['(최우선)매도호가'])))
            bid_price = abs(int(self.dynamicCall("GetCommRealData(QString, int)",
                                         s_code, self.realType.REAL_TYPE[real_type]['(최우선)매수호가'])))
            trad_volume = abs(int(self.dynamicCall("GetCommRealData(QString, int)",
                                         s_code, self.realType.REAL_TYPE[real_type]['거래량'])))
            c_trad_volume = abs(int(self.dynamicCall("GetCommRealData(QString, int)",
                                         s_code, self.realType.REAL_TYPE[real_type]['누적거래량'])))
            high_price = abs(int(self.dynamicCall("GetCommRealData(QString, int)",
                                         s_code, self.realType.REAL_TYPE[real_type]['고가'])))
            opening_price = abs(int(self.dynamicCall("GetCommRealData(QString, int)",
                                         s_code, self.realType.REAL_TYPE[real_type]['시가'])))
            low_price = abs(int(self.dynamicCall("GetCommRealData(QString, int)",
                                         s_code, self.realType.REAL_TYPE[real_type]['저가'])))

            if s_code in self.candidates_dict:
                self.candidates_dict[s_code].update({"체결시간": concluded_time})
                self.candidates_dict[s_code].update({"현재가": current_price})
                self.candidates_dict[s_code].update({"전일대비": range_day2day})
                self.candidates_dict[s_code].update({"등락율": today_range})
                self.candidates_dict[s_code].update({"(최우선)매도호가": ask_price})
                self.candidates_dict[s_code].update({"(최우선)매수호가": bid_price})
                self.candidates_dict[s_code].update({"거래량": trad_volume})
                self.candidates_dict[s_code].update({"누적거래량": c_trad_volume})
                self.candidates_dict[s_code].update({"고가": high_price})
                self.candidates_dict[s_code].update({"시가": opening_price})
                self.candidates_dict[s_code].update({"저가": low_price})

            now = datetime.now().strftime("%H:%M:%S")
            if now >= self.TIME_TO_SUMMARIZE and self.is_summarized:
                # if bool(self.failed_dict):
                #     f1 = open(self.FAILED_PATH, "a", encoding="utf8")
                #     f1.write(json.dumps(self.failed_dict, indent=4, ensure_ascii=False))
                #     f1.close()
                # if bool(self.targets_dict):
                #     f2 = open(self.TARGETS_PATH, "a", encoding="utf8")
                #     temp_dict = copy.deepcopy(self.targets_dict)
                #     for c in temp_dict:
                #         if "last_bought_time" in temp_dict[c]:
                #             temp_dict[c].update({"last_bought_time": temp_dict[c]["last_bought_time"].strftime("%H:%M:%S")})
                #         if "last_sold_time" in temp_dict[c]:
                #             temp_dict[c].update({"last_sold_time": temp_dict[c]["last_sold_time"].strftime("%H:%M:%S")})
                #         if "cut_time" in temp_dict[c]:
                #             temp_dict[c].update({"cut_time": temp_dict[c]["cut_time"].strftime("%H:%M:%S")})
                #     f2.write(json.dumps(temp_dict, indent=4, ensure_ascii=False))
                #     f2.close()
                self.is_summarized = False

            # 오늘 한번이라도 체결된 경우
            if s_code in self.targets_dict.keys():
                target = self.targets_dict[s_code]
                if target['매입단가'] > 0 and target['주문가능수량'] > 0 and self.targets_dict[s_code]["holding"]:
                    if now < self.TIME_TO_STOP_TRADING:
                        # sellable = False
                        ror = (bid_price - target['매입단가']) / target['매입단가'] * 100
                        target.update({"현재수익률": ror})
                        # target.update({"initial_price": target["매입단가"]})
                        if ror > target["최고수익률"]:
                            target.update({"최고수익률": ror})
                            target.update({"최고가": bid_price})
                        if ror < target["최저수익률"]:
                            target.update({"최저수익률": ror})

                        gap = 0
                        if target["최고수익률"] >= 0.3:
                            gap = target["최고수익률"] - ror

                        # after_3minute = target["last_bought_time"] + timedelta(minutes=3)
                        # after_5minute = target["last_bought_time"] + timedelta(minutes=5)
                        # if ror < 1 and after_3minute.minute <= datetime.now().minute < after_5minute.minute\
                        #         and not any(etf in self.targets_dict[s_code]["종목명"] for etf in self.ETF_ALIAS):
                        #     self.dynamicCall(
                        #         "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                        #         ["신규매도", self.candidates_dict[s_code]["주문용화면번호"],
                        #          self.account_num, 2, s_code, target['주문가능수량'], 0,
                        #          self.realType.SEND_TYPE['거래구분']['시장가'], ""])
                        #     target.update({"last_sold_time": datetime.now()})
                        #     target.update({"last_sold_price": bid_price})
                        #     target.update({"result": "기대이하"})
                        #     target.update({"continuous_cut": False})
                        # 최고수익률 대비 익절 및 스탑로스
                        if gap > 0:
                            rod = 0.5
                            result = "스탑로스"
                            if ror > 3:
                                rod = 0.3
                                result = "익절"
                            if gap >= rod:
                                if gap > 0.7:
                                    self.dynamicCall(
                                        "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                        ["신규매도", self.candidates_dict[s_code]["주문용화면번호"],
                                         self.account_num, 2, s_code, target['주문가능수량'], 0,
                                         self.realType.SEND_TYPE['거래구분']['시장가'], ""])
                                    target.update({"last_sold_time": datetime.now()})
                                    target.update({"last_sold_price": bid_price})
                                    target.update({"result": result})
                                    target.update({"continuous_cut": False})
                                else:
                                    if "cut_time" not in target:
                                        target.update({"cut_time": datetime.now()})
                                        target.update({"continuous_cut": False})
                                    else:
                                        after_1minute = target["cut_time"] + timedelta(minutes=1)
                                        after_2minute = target["cut_time"] + timedelta(minutes=2)
                                        if datetime.now().minute == after_1minute.minute:
                                            target.update({"continuous_cut": True})
                                        elif target["continuous_cut"] and datetime.now().minute == after_2minute.minute:
                                            self.dynamicCall(
                                                "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                                ["신규매도", self.candidates_dict[s_code]["주문용화면번호"],
                                                 self.account_num, 2, s_code, target['주문가능수량'], 0,
                                                 self.realType.SEND_TYPE['거래구분']['시장가'], ""])
                                            target.update({"last_sold_time": datetime.now()})
                                            target.update({"last_sold_price": bid_price})
                                            target.update({"result": result})
                                            target.update({"continuous_cut": False})
                                        elif datetime.now().minute > after_2minute.minute:
                                            target.update({"cut_time": datetime.now()})
                                            target.update({"continuous_cut": False})
                            else:
                                if "cut_time" in target:
                                    after_1minute = target["cut_time"] + timedelta(minutes=1)
                                    if datetime.now().minute == after_1minute.minute:
                                        target.update({"continuous_cut": False})
                        # 일반 손절
                        elif gap <= 0:
                            if ror < -0.7:
                                self.dynamicCall(
                                    "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                    ["신규매도", self.candidates_dict[s_code]["주문용화면번호"],
                                     self.account_num, 2, s_code, target['주문가능수량'], 0,
                                     self.realType.SEND_TYPE['거래구분']['시장가'], ""])
                                target.update({"last_sold_time": datetime.now()})
                                target.update({"last_sold_price": bid_price})
                                target.update({"result": "손절"})
                                target.update({"continuous_cut": False})
                            elif ror < -0.4:
                                if "cut_time" not in target:
                                    target.update({"cut_time": datetime.now()})
                                    target.update({"continuous_cut": False})
                                else:
                                    after_1minute = target["cut_time"] + timedelta(minutes=1)
                                    after_2minute = target["cut_time"] + timedelta(minutes=2)
                                    if datetime.now().minute == after_1minute.minute:
                                        target.update({"continuous_cut": True})
                                    elif target["continuous_cut"] and datetime.now().minute == after_2minute.minute:
                                        self.dynamicCall(
                                            "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                            ["신규매도", self.candidates_dict[s_code]["주문용화면번호"],
                                             self.account_num, 2, s_code, target['주문가능수량'], 0,
                                             self.realType.SEND_TYPE['거래구분']['시장가'], ""])
                                        target.update({"last_sold_time": datetime.now()})
                                        target.update({"last_sold_price": bid_price})
                                        target.update({"result": "손절"})

                                        target.update({"continuous_cut": False})
                                    elif datetime.now().minute > after_2minute.minute:
                                        target.update({"cut_time": datetime.now()})
                                        target.update({"continuous_cut": False})
                            else:
                                if "cut_time" in target:
                                    after_1minute = target["cut_time"] + timedelta(minutes=1)
                                    if datetime.now().minute == after_1minute.minute:
                                        target.update({"continuous_cut": False})
                    else:
                        # 15시 18분 이후 모두 매도
                        self.dynamicCall(
                            "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                            ["신규매도", self.candidates_dict[s_code]["주문용화면번호"], self.account_num, 2, s_code,
                             target['주문가능수량'],
                             0, self.realType.SEND_TYPE['거래구분']['시장가'], ""])
                        target.update({"last_sold_time": datetime.now()})
                        target.update({"last_sold_price": bid_price})
                        target.update({"result": "장종료"})
                else:
                    if now < self.TIME_TO_STOP_TRADING:
                        if "last_sold_price" in target:
                            ror_tracked = (current_price - target['last_sold_price']) / target['last_sold_price'] * 100
                            if ror_tracked < target["추적수익률"]:
                                target.update({"추적수익률": ror_tracked})


            # 기존 잔고에 존재했던 종목인 경우
            elif s_code in self.holdings_dict.keys():
                asd = self.holdings_dict[s_code]
                self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                 ["신규매도", self.candidates_dict[s_code]["주문용화면번호"],
                                  self.account_num, 2, s_code, asd['매매가능수량'], 0,
                                  self.realType.SEND_TYPE['거래구분']['시장가'], ""])

            # 현재 매매중이 아닌 경우
            if s_code in self.candidates_dict and s_code not in self.on_trading_list \
                    and s_code not in self.holdings_dict.keys():

                if datetime.now().strftime("%H:%M:%S") > self.TIME_TO_START_TO_BUY and ask_price > 0:
                    buyable = False
                    quantity = int(self.BUYING_AMOUNT / ask_price)

                    # 매매한 적이 한번이라도 있는 경우
                    if s_code in self.targets_dict and not self.targets_dict[s_code]["holding"] \
                            and "last_sold_time" in self.targets_dict[s_code]:
                        wait = datetime.now() - self.targets_dict[s_code]["last_sold_time"]
                        if wait.seconds > self.MINUTES_5:
                            buyable = True
                    elif s_code not in self.targets_dict:
                        buyable = True

                    if buyable:
                        # if quantity > 0:
                        # 지정가 매수
                        # order_success = self.dynamicCall(
                        #     "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                        #     ["신규매수", self.candidates_dict[s_code]["주문용화면번호"],
                        #      self.account_num, 1, s_code, quantity, ask_price,
                        #      self.realType.SEND_TYPE['거래구분']['지정가'], ""])
                        # 시장가 매수
                        order_success = self.dynamicCall(
                            "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                            ["신규매수", self.candidates_dict[s_code]["주문용화면번호"],
                             self.account_num, 1, s_code, quantity, 0,
                             self.realType.SEND_TYPE['거래구분']['시장가'], ""])

                        if order_success != 0:
                            if s_code in self.failed_dict:
                                self.failed_dict[s_code].append(datetime.now().strftime("%H:%M:%S"))
                            else:
                                self.failed_dict.update({
                                    s_code: [datetime.now().strftime("%H:%M:%S")]
                                })
                        if s_code not in self.on_trading_list:
                            self.on_trading_list.append(s_code)

            not_meme_list = list(self.unfulfilled_dict)
            for order_num in not_meme_list:
                code = self.unfulfilled_dict[order_num]["종목코드"]
                order_price = self.unfulfilled_dict[order_num]['주문가격']
                not_quantity = self.unfulfilled_dict[order_num]['미체결수량']
                order_cat = self.unfulfilled_dict[order_num]['주문구분']
                stock_name = self.unfulfilled_dict[order_num]['종목명']
                origin_order_number = "Empty"
                if "원주문번호" in self.unfulfilled_dict[order_num]:
                    origin_order_number = self.unfulfilled_dict[order_num]['원주문번호']

                #  and ask_price > order_price
                if order_cat == "매수" and not_quantity > 0:
                    wait = datetime.now() - self.unfulfilled_dict[order_num]["order_time"]
                    if s_code in self.candidates_dict and wait.seconds >= self.MINUTES_2:
                        order_success = self.dynamicCall(
                            "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                            ["매수취소", self.candidates_dict[s_code]["주문용화면번호"], self.account_num, 3, code, 0, 0,
                             self.realType.SEND_TYPE['거래구분']['지정가'], order_num]
                        )

                        # if order_success == 0:
                        #     logging.info("[%s] 매수 취소 전달 성공", _name)
                        # else:
                        #     logging.info("[%s] 매수 취소 전달 실패", _name)

                elif order_cat == "매수" and not_quantity == 0:
                    logging.info("[주문번호: %s] %s Buy", order_num, stock_name)
                    del self.unfulfilled_dict[order_num]

                elif order_cat == "매도" and not_quantity == 0:
                    logging.info("[주문번호: %s] %s Sell", order_num, stock_name)
                    del self.unfulfilled_dict[order_num]

                elif order_cat == "매수취소" and not_quantity == 0:
                    logging.info("[원주문번호: %s] %s Cancel to buy", origin_order_number, stock_name)
                    if code in self.on_trading_list and code not in self.targets_dict:
                        self.on_trading_list.remove(code)
                    del self.unfulfilled_dict[order_num]

                elif order_cat == "매도취소" and not_quantity == 0:
                    logging.info("[원주문번호: %s] %s Cancel to sell", origin_order_number, stock_name)
                    del self.unfulfilled_dict[order_num]

    def _chejan_handler(self, sGubun, nItemCnt, sFidList):
        if int(sGubun) == 0:  # 주문체결
            s_code = self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['주문체결']['종목코드'])[1:]
            stock_name = self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['주문체결']['종목명']).strip()
            origin_order_number = self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['주문체결']['원주문번호'])
            order_number = self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['주문체결']['주문번호'])
            order_status = self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['주문체결']['주문상태'])
            order_quan = int(self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['주문체결']['주문수량']))
            order_price = int(self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['주문체결']['주문가격']))
            not_chegual_quan = int(self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['주문체결']['미체결수량']))
            order_cat = self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['주문체결']['주문구분']) \
                .strip().lstrip('+').lstrip('-')
            chegual_time_str = self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['주문체결']['주문/체결시간'])
            _chegual_price = self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['주문체결']['체결가'])
            chegual_price = 0 if _chegual_price == '' else int(_chegual_price)
            _chegual_quantity = self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['주문체결']['체결량'])
            chegual_quantity = 0 if _chegual_quantity == '' else int(_chegual_quantity)
            current_price = abs(int(self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['주문체결']['현재가'])))
            first_sell_price = abs(int(self.dynamicCall("GetChejanData(int)",
                                                        self.realType.REAL_TYPE['주문체결']['(최우선)매도호가'])))
            first_buy_price = abs(int(self.dynamicCall("GetChejanData(int)",
                                                       self.realType.REAL_TYPE['주문체결']['(최우선)매수호가'])))

            # 새로 들어온 주문이면 주문번호 할당
            if order_number not in self.unfulfilled_dict.keys():
                self.unfulfilled_dict.update({order_number: {"order_time": datetime.now()}})

            self.unfulfilled_dict[order_number].update({"종목코드": s_code})
            self.unfulfilled_dict[order_number].update({"주문번호": order_number})
            self.unfulfilled_dict[order_number].update({"종목명": stock_name})
            self.unfulfilled_dict[order_number].update({"주문상태": order_status})
            self.unfulfilled_dict[order_number].update({"주문수량": order_quan})
            self.unfulfilled_dict[order_number].update({"주문가격": order_price})
            self.unfulfilled_dict[order_number].update({"미체결수량": not_chegual_quan})
            self.unfulfilled_dict[order_number].update({"원주문번호": origin_order_number})
            self.unfulfilled_dict[order_number].update({"주문구분": order_cat})
            self.unfulfilled_dict[order_number].update({"주문/체결시간": chegual_time_str})
            self.unfulfilled_dict[order_number].update({"체결가": chegual_price})
            self.unfulfilled_dict[order_number].update({"체결량": chegual_quantity})
            self.unfulfilled_dict[order_number].update({"현재가": current_price})
            self.unfulfilled_dict[order_number].update({"(최우선)매도호가": first_sell_price})
            self.unfulfilled_dict[order_number].update({"(최우선)매수호가": first_buy_price})

        elif int(sGubun) == 1:  # 잔고
            s_code = self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['잔고']['종목코드'])[1:]
            stock_name = self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['잔고']['종목명']).strip()
            current_price = abs(int(self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['잔고']['현재가'])))
            stock_quan = int(self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['잔고']['보유수량']))
            like_quan = int(self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['잔고']['주문가능수량']))
            buy_price = abs(int(self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['잔고']['매입단가'])))
            total_buy_price = int(self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['잔고']['총매입가']))
            meme_gubun = self.realType.REAL_TYPE['매도수구분'][
                self.dynamicCall("GetChejanData(int)", self.realType.REAL_TYPE['잔고']['매도매수구분'])]
            first_sell_price = abs(int(self.dynamicCall("GetChejanData(int)",
                                                        self.realType.REAL_TYPE['잔고']['(최우선)매도호가'])))
            first_buy_price = abs(int(self.dynamicCall("GetChejanData(int)",
                                                       self.realType.REAL_TYPE['잔고']['(최우선)매수호가'])))

            # 초기 값만 세팅하고, 거래 이벤트에서만 업데이트될 컬럼들
            if s_code not in self.targets_dict.keys():
                self.targets_dict.update({
                    s_code: {
                        "매매횟수": 0, "최고수익률": 0, "최저수익률": 0, "현재수익률": 0, "추적수익률": 0, "최고가": buy_price,
                        "last_bought_time": datetime.now(), "last_bought_price": buy_price
                    }})

            # 체결 이후 업데이트될 수 있는 컬럼들
            self.targets_dict[s_code].update({"현재가": current_price})
            self.targets_dict[s_code].update({"종목코드": s_code})
            self.targets_dict[s_code].update({"종목명": stock_name})
            self.targets_dict[s_code].update({"보유수량": stock_quan})
            self.targets_dict[s_code].update({"주문가능수량": like_quan})
            self.targets_dict[s_code].update({"매입단가": buy_price})
            self.targets_dict[s_code].update({"총매입가": total_buy_price})
            self.targets_dict[s_code].update({"매도매수구분": meme_gubun})
            self.targets_dict[s_code].update({"(최우선)매도호가": first_sell_price})
            self.targets_dict[s_code].update({"(최우선)매수호가": first_buy_price})
            if stock_quan > 0:
                self.targets_dict[s_code].update({"holding": True})

            # 한번이라도 주문했고, 잔고에 남아있지 않으면 삭제함.
            if s_code in self.on_trading_list and self.targets_dict[s_code]["총매입가"] < 1:
                self.on_trading_list.remove(s_code)
                target = self.targets_dict[s_code]
                target.update({"매매횟수": self.targets_dict[s_code]["매매횟수"] + 1})
                logging.info("매매 결과: \n%s", pformat(target))
                f = open(self.LOG_PATH, "a", encoding="utf8")
                f.write("%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n" %
                        (datetime.now().strftime("%H:%M:%S"),
                         target["종목명"],
                         target["last_bought_price"],
                         target["last_bought_time"].strftime("%H:%M:%S"),
                         target["last_sold_price"],
                         target["last_sold_time"].strftime("%H:%M:%S"),
                         target["result"],
                         round(target["최고수익률"], 6),
                         round(target["최저수익률"], 6),
                         round(target["현재수익률"], 6)))
                f.close()

                target.update({"최고수익률": 0})
                target.update({"최저수익률": 0})
                target.update({"현재수익률": 0})
                target.update({"추적수익률": 0})

                self.dynamicCall("SetRealRemove(QString, QString)", self.candidates_dict[s_code]['화면번호'], s_code)
                self.targets_dict[s_code].update({"holding": False})
                del self.candidates_dict[s_code]

    def today_log(self):
        f = open(self.LOG_PATH, "a", encoding="utf8")
        f.write("___시간__\t__종목명__\t매수가\t_매수시간_\t매도가\t_매도시간_\t_결과_\t최고수익률\t최저수익률\t현재수익률\n")
        f.close()
