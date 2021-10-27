
"""
Default_API.py
매수 및 매도 함수 구현
"""
import requests
import json
import os, sys, ctypes
import win32com.client
import pandas as pd
from datetime import datetime
import time, calendar
from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
import json


cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')
cpOhlc = win32com.client.Dispatch("CpSysDib.StockChart")

def check_creon_system():
    """
    Creon Systen Check
    """
    if not ctypes.windll.shell32.IsUserAnAdmin():
        print('check_creon_system() : admin user -> Failed')
        return False
    if (cpStatus.IsConnect == 0):
        print('check_creon_system() : connect to server -> Failed')
        return False
    if (cpTradeUtil.TradeInit(0) != 0):
        print('check_creon_system(): init trade -> Failed')
        return False
    return True


def post_message(slack_data):
    """
    :param slack_data
    :return:
    특정 채널에 slack message를 보냄
    * 채널은 webhook_url로 선택
    * Default 채널: creon-logs-2
    """
    webhook_url = "***********************************"
    response = requests.post(webhook_url, data=json.dumps(slack_data),
                             headers={'Content-Type': 'application/json'})
    if response.status_code != 200:
        raise ValueError(
            'Request to slack returned an error %s, the response is:\n%s'
            % (response.status_code, response.text))


def dbgout(message):
    """
    :param message
    :return:
    message를 cmd 및 slack에 표시합니다.
    """
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
    strbuf = datetime.now().strftime('[%m/%d %H:%M:%S]') + message
    slack_data = {
        "blocks": [
            {
                "type": "section",
                "text": {
                    "type": "plain_text",
                    "text": strbuf,
                    "emoji": True
                }
            }
        ]
    }
    post_message(slack_data)


def printlog(message, *args):
    """
    :param message:
    :param args:
    :return:
    message 와 args를 cmd에 표시합니다
    """
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message, *args)


def get_current_price(code):
    """
    :param code: 종목 코드
    :return:
    현재가, 매도호가, 매수호가
    """
    cpStock.SetInputValue(0, code)
    cpStock.BlockRequest()
    item = {}
    item['cur_price'] = cpStock.GetHeaderValue(11)
    item['ask'] = cpStock.GetHeaderValue(16)
    item['bid'] = cpStock.GetHeaderValue(17)
    return item['cur_price'], item['ask'], item['bid']


def get_ohlc(code, qty):
    """
    :param code: 종목 코드
    :param qty: 요청 데이터 개수 (ohlcv 날짜의 개수)
    :return:
    dataframe

    <참고>
    https://money2.creontrade.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=102&page=1&searchString=CpSysDib.StockChart&p=8841&v=8643&m=9505
    """
    cpOhlc.SetInputValue(0, code)
    cpOhlc.SetInputValue(1, ord('2'))
    cpOhlc.SetInputValue(4, qty)
    cpOhlc.SetInputValue(5, [0, 2, 3, 4, 5])
    cpOhlc.SetInputValue(6, ord('D'))
    cpOhlc.SetInputValue(9, ord('1'))
    cpOhlc.BlockRequest()
    count = cpOhlc.GetHeaderValue(3)
    columns = ['open', 'high', 'low', 'close']
    index = []
    rows = []
    for i in range(count):
        index.append(cpOhlc.GetDataValue(0, i))
        rows.append([cpOhlc.GetDataValue(1, i), cpOhlc.GetDataValue(2, i),
                     cpOhlc.GetDataValue(3, i), cpOhlc.GetDataValue(4, i)])
    df = pd.DataFrame(rows, columns=columns, index=index)
    return df


def get_stock_balance(code, verbose=True):
    """
    :param code: 종목 코드 혹은 'ALL'
    :param verbose: log 출력 여부
    :return:
    1) 단일 종목
    종목명과 보유 수량을 반환
    2) ALL
    코드, 종목명, 보유 수량을 담은 리스트를 반환
    [{'code': -, 'name': -, 'qty'}, ...]

    <참고>
    https://money2.creontrade.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=176&page=1&searchString=CpTrade.CpTd6033&p=8841&v=8643&m=9505
    """
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체, 1:주식, 2:선물/옵션
    cpBalance.SetInputValue(0, acc)         # 계좌번호
    cpBalance.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    cpBalance.SetInputValue(2, 50)          # 요청 건수(최대 50)
    cpBalance.BlockRequest()
    if code == 'ALL' and verbose:
        dbgout('계좌명: ' + str(cpBalance.GetHeaderValue(0)))
        dbgout('결제잔고수량 : ' + str(cpBalance.GetHeaderValue(1)))
        dbgout('평가금액: ' + str(cpBalance.GetHeaderValue(3)))
        dbgout('평가손익: ' + str(cpBalance.GetHeaderValue(4)))
        dbgout('종목수: ' + str(cpBalance.GetHeaderValue(7)))
    stocks = []
    for i in range(cpBalance.GetHeaderValue(7)):
        stock_code = cpBalance.GetDataValue(12, i)  # 종목코드
        stock_name = cpBalance.GetDataValue(0, i)   # 종목명
        stock_qty = cpBalance.GetDataValue(15, i)   # 수량
        if code == 'ALL':
            if verbose:
                dbgout(str( i +1) + ' ' + stock_code + '(' + stock_name + ')'
                       + ':' + str(stock_qty))
            stocks.append({'code': stock_code, 'name': stock_name,
                           'qty': stock_qty})
        if stock_code == code:
            return stock_name, stock_qty
    if code == 'ALL':
        return stocks
    else:
        stock_name = cpCodeMgr.CodeToName(code)
        return stock_name, 0


def get_stock_status(code):
    """
    :param code: 종목 코드
    :return:
    종목명 (str), 보유 수량 (int), 수익률 (double)
    <참고>
    https://money2.creontrade.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=176&page=1&searchString=CpTrade.CpTd6033&p=8841&v=8643&m=9505
    """
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체, 1:주식, 2:선물/옵션
    cpBalance.SetInputValue(0, acc)         # 계좌번호
    cpBalance.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    cpBalance.SetInputValue(2, 50)          # 요청 건수(최대 50)
    cpBalance.BlockRequest()
    for i in range(cpBalance.GetHeaderValue(7)):
        stock_code = cpBalance.GetDataValue(12, i)      # 종목코드
        stock_name = cpBalance.GetDataValue(0, i)       # 종목명
        stock_qty = cpBalance.GetDataValue(15, i)       # 수량
        stock_yield = cpBalance.GetDataValue(11, i)     # 수익률
        if stock_code == code:
            return stock_name, stock_qty, stock_yield
    stock_name = cpCodeMgr.CodeToName(code)
    return stock_name, 0, 0.0


def get_stock_list():
    """
    :return:
    현재 보유 중인 종목 리스트를
    1) have_stock_list.json 파일로 저장
    2) 반환
    """
    ret_stock_list = []
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]          # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1)     # -1:전체, 1:주식, 2:선물/옵션
    cpBalance.SetInputValue(0, acc)             # 계좌번호
    cpBalance.SetInputValue(1, accFlag[0])      # 상품구분 - 주식 상품 중 첫번째
    cpBalance.SetInputValue(2, 50)              # 요청 건수(최대 50)
    cpBalance.BlockRequest()
    for i in range(cpBalance.GetHeaderValue(7)):
        stock_code = cpBalance.GetDataValue(12, i)
        ret_stock_list.append(stock_code)
    with open('have_stock_list.json', 'w') as f:
        json.dump(ret_stock_list, f)
    return ret_stock_list


def stock_monitor(code_list, upper_yield_limit, lower_yield_limit):
    """
    :param code_list: 모니터할 종목 리스트
    :param upper_yield_limit: 수익률 상한 (e.g. 50 -> +50%)
    :param lower_yield_limit: 수익률 하한 (e.g. 50 -> -50%)
    :return:
    상한/하한 수익률을 돌파한 종목 리스트를 반환
    """
    ret_code_list = []
    for code in code_list:
        name, qty, _yield = get_stock_status(code)
        if _yield >= upper_yield_limit or _yield <= -lower_yield_limit:
            ret_code_list.append(code)
    return ret_code_list


def get_current_cash():
    """
    :return: 주문 가능 금액 (증거금 100%)
    """
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]    # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체, 1:주식, 2:선물/옵션
    cpCash.SetInputValue(0, acc)              # 계좌번호
    cpCash.SetInputValue(1, accFlag[0])      # 상품구분 - 주식 상품 중 첫번째
    cpCash.BlockRequest()
    return cpCash.GetHeaderValue(9) # 증거금 100% 주문 가능 금액


def buy_stock(code, buy_qty):
    """
    :param code: 종목 코드
    :param buy_qty: 매수 수량
    :return:
    최우선 FOK로 매수
    """
    try:
        time_now = datetime.now()
        current_price, ask_price, bid_price = get_current_price(code)
        stock_name, stock_qty = get_stock_balance(code)  # 종목명과 보유수량 조회
        # 기 매수 종목
        if stock_qty > 0:
            return True
        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
        accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체,1:주식,2:선물/옵션
        cpOrder.SetInputValue(0, "2")           # 1:매도, 2:매수
        cpOrder.SetInputValue(1, acc)           # 계좌번호
        cpOrder.SetInputValue(2, accFlag[0])    # 상품구분 - 주식 상품 중 첫번째
        cpOrder.SetInputValue(3, code)          # 종목코드
        cpOrder.SetInputValue(4, buy_qty)       # 매수할 수량
        cpOrder.SetInputValue(5, ask_price)     # 매수 희망 가격
        cpOrder.SetInputValue(7, "0")           # 주문조건 0:기본, 1:IOC, 2:FOK
        cpOrder.SetInputValue(8, "01")          # 주문호가 01:보통, 03:시장가
        # 05:조건부, 12:최유리, 13:최우선
        ret = cpOrder.BlockRequest()
        dbgout('* 매도호가 기본 매수: ' + str(stock_name) + ', ' + str(code) + ', ' +
               str(buy_qty) + ' (주문 ret: ' + str(ret) + ')')
        if ret == 1 or ret == 2:
            dbgout('주문 오류.')
            return False
        if ret == 4:
            remain_time = cpStatus.LimitRequestRemainTime
            dbgout('주의: 연속 주문 제한에 걸림. 대기 시간:' + str(remain_time/1000))
            time.sleep(remain_time /1000)
            return False
        rqStatus = cpOrder.GetDibStatus()
        errMsg = cpOrder.GetDibMsg1()
        if rqStatus != 0:
            printlog("주문 실패: ", rqStatus, errMsg)
        time.sleep(3)
        stock_name, bought_qty = get_stock_balance(code)
        if bought_qty > 0:
            dbgout("<" + str(stock_name) + ' , ' + str(code) +
                   "> : " + str(bought_qty) + "EA 매수 완료")
            return True
        return False

    except Exception as ex:
        dbgout("매수 함수 에러 발생 " + "(에러 내용: " + str(ex) + ")")
        return False


def buy_stock_list(code_list):
    """
    :param code_list: 매수 종목 코드 리스트
    :return:
    code_list 동일 가중 포트폴리오 매수
    """
    target_buy_count = 30                       # 매수할 종목 수
    buy_percent = 1 / target_buy_count * 0.95
    total_cash = int(get_current_cash())        # 100% 증거금 주문 가능 금액 조회
    buy_amount = total_cash * buy_percent       # 종목별 주문 금액 계산
    for code in code_list:
        while True:
            time.sleep(5)
            current_price, ask_price, bid_price = get_current_price(code)
            buy_qty = buy_amount // bid_price
            buy_flag = buy_stock(code, buy_qty)
            if buy_flag:
                break
    dbgout('=========================')
    dbgout('종목 리스트 매수 완료')
    get_stock_balance('ALL')
    dbgout('=========================')


def sell_stock(code):
    """
    :param code: 종목 코드
    :return:
    최유리 IOC로 전량 매도
    """
    try:
        time_now = datetime.now()
        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]          # 계좌번호
        accFlag = cpTradeUtil.GoodsList(acc, 1)     # -1:전체,1:주식,2:선물/옵션
        while True:
            t_now = datetime.now()
            t_exit = t_now.replace(hour=15, minute=20, second=0, microsecond=0)
            if t_now > t_exit:
                return False
            stock_name, stock_qty = get_stock_balance(code=code, verbose=False)
            if stock_qty == 0:
                dbgout("<" + str(stock_name) + ' , ' + str(code) +
                       "> 매도 완료")
                return True
            current_price, ask_price, bid_price = get_current_price(code)
            cpOrder.SetInputValue(0, "1")               # 1:매도, 2:매수
            cpOrder.SetInputValue(1, acc)               # 계좌번호
            cpOrder.SetInputValue(2, accFlag[0])        # 상품구분 - 주식 상품 중 첫번째
            cpOrder.SetInputValue(3, code)              # 종목코드
            cpOrder.SetInputValue(4, stock_qty)         # 매도할 수량
            # cpOrder.SetInputValue(5, bid_price)         # 매도 희망 가격
            cpOrder.SetInputValue(7, "1")               # 주문조건 0:기본, 1:IOC, 2:FOK
            cpOrder.SetInputValue(8, "12")              # 주문호가 01:보통, 03:시장가
            # 05:조건부, 12:최유리, 13:최우선
            ret = cpOrder.BlockRequest()
            dbgout('* 최유리 IOC 매도: ' + str(stock_name) + ', ' + str(code) + ', ' +
                   str(stock_qty) + ' (주문 ret: ' + str(ret) + ')')
            if ret == 1 or ret == 2:
                dbgout('주문 오류.')
            if ret == 4:
                remain_time = cpStatus.LimitRequestRemainTime
                dbgout('주의: 연속 주문 제한에 걸림. 대기 시간:' + str(remain_time /1000))
                time.sleep(remain_time / 1000)
            rqStatus = cpOrder.GetDibStatus()
            errMsg = cpOrder.GetDibMsg1()
            if rqStatus != 0:
                printlog("주문 실패: ", rqStatus, errMsg)
            time.sleep(30)
    except Exception as ex:
        dbgout("매도 함수 에러 발생 " + "(에러 내용: " + str(ex) + ")")
        return False


def sell_stock_list(code_list):
    """
    :param code_list: 매도 종목 코드 리스트
    :return:
    code_list 포트폴리오 매도
    """
    for code in code_list:
        sell_stock(code)
    dbgout('========================')
    dbgout('종목 리스트 매도 완료')
    dbgout('========================')
    get_stock_balance('ALL')


def sell_stock_all():
    """
    :return:
    보유한 모든 종목을 최유리 지정가 IOC 조건으로 매도
    """
    try:
        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
        accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션
        while True:
            t_now = datetime.now()
            t_exit = t_now.replace(hour=15, minute=20, second=0, microsecond=0)
            if t_now > t_exit:
                return False
            stocks = get_stock_balance('ALL')
            total_qty = 0
            for s in stocks:
                total_qty += s['qty']
            if total_qty == 0:
                return True
            for s in stocks:
                if s['qty'] != 0:
                    cpOrder.SetInputValue(0, "1")           # 1:매도, 2:매수
                    cpOrder.SetInputValue(1, acc)           # 계좌번호
                    cpOrder.SetInputValue(2, accFlag[0])    # 주식상품 중 첫번째
                    cpOrder.SetInputValue(3, s['code'])     # 종목코드
                    cpOrder.SetInputValue(4, s['qty'])      # 매도수량
                    cpOrder.SetInputValue(7, "1")           # 조건 0:기본, 1:IOC, 2:FOK
                    cpOrder.SetInputValue(8, "12")          # 호가 12:최유리, 13:최우선
                                                            # 최유리 IOC 매도 주문 요청
                    ret = cpOrder.BlockRequest()
                    dbgout('* 최우선 기본 매도: ' + str(s['code']) + ', ' + str(s['qty']) +
                           ' (주문 ret: ' + str(ret) + ')')
                    if ret == 1 or ret == 2:
                        dbgout('주문 오류.')
                    if ret == 4:
                        remain_time = cpStatus.LimitRequestRemainTime
                        dbgout('주의: 연속 주문 제한에 걸림. 대기 시간:' + str(remain_time/1000))
                        time.sleep(remain_time / 1000)
                    rqStatus = cpOrder.GetDibStatus()
                    errMsg = cpOrder.GetDibMsg1()
                    if rqStatus != 0:
                        printlog("주문 실패: ", rqStatus, errMsg)
                time.sleep(3)
            time.sleep(30)
        dbgout('=========================')
        dbgout('모든 종목 매도 완료')
        dbgout('=========================')
        get_stock_balance('ALL')

    except Exception as ex:
        dbgout("매도 함수 에러 발생 " + "(에러 내용: " + str(ex) + ")")


"""
Execution
"""
if __name__ == '__main__':
    try:
        printlog('check_creon_system(): ', check_creon_system())  # 크레온 접속 점검
        stocks = get_stock_balance('ALL')  # 보유한 모든 종목 조회
        total_cash = int(get_current_cash())  # 100% 증거금 주문 가능 금액 조회
        dbgout('100% 증거금 주문 가능 금액: ' + str(total_cash))
        dbgout('현재 시각: ' + datetime.now().strftime('%m/%d %H:%M:%S'))
        today = datetime.today().weekday()
        if today == 5 or today == 6:  # 토요일이나 일요일이면 자동 종료
            printlog('주말 자동 종료')
            sys.exit(0)

        # buy_stock('A069500', 1)
        # sell_stock('A069500')

        sys.exit(0)
    except Exception as ex:
        dbgout("Main 함수 에러 발생 " + "(에러 내용: " + str(ex) + ")")



