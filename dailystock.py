from slacker import Slacker

slack = Slacker('xoxb-1600375133346-1606657440468-Q20yu2dp8KwWD3Qu9ZicH7DX')

import win32com.client
from datetime import datetime


# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()
 
# 현재가 객체 구하기
objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
objStockMst.SetInputValue(0, 'A005930')   #종목 코드 - 삼성전자
objStockMst.BlockRequest()
 
# 현재가 통신 및 통신 에러 처리 
rqStatus = objStockMst.GetDibStatus()
rqRet = objStockMst.GetDibMsg1()
print("통신상태", rqStatus, rqRet)
if rqStatus != 0:
    exit()


#시간 나오기
today=datetime.today()
#hour=datetime.today().hour
#현재시간메세지보내기
slack.chat.post_message('#stock_daily', "오늘날짜 : " + str(today))
#slack.chat.post_message('#stock_daily', "현재시각 : " + str(hour))

  
# 현재가 정보 조회
code = objStockMst.GetHeaderValue(0)  #종목코드
name= objStockMst.GetHeaderValue(1)  # 종목명
time= objStockMst.GetHeaderValue(4)  # 시간
cprice= objStockMst.GetHeaderValue(11) # 종가
diff= objStockMst.GetHeaderValue(12)  # 대비
open= objStockMst.GetHeaderValue(13)  # 시가
high= objStockMst.GetHeaderValue(14)  # 고가
low= objStockMst.GetHeaderValue(15)   # 저가
offer = objStockMst.GetHeaderValue(16)  #매도호가
bid = objStockMst.GetHeaderValue(17)   #매수호가
vol= objStockMst.GetHeaderValue(18)   #거래량
vol_value= objStockMst.GetHeaderValue(19)  #거래대금
 
# 예상 체결관련 정보
exFlag = objStockMst.GetHeaderValue(58) #예상체결가 구분 플래그
exPrice = objStockMst.GetHeaderValue(55) #예상체결가
exDiff = objStockMst.GetHeaderValue(56) #예상체결가 전일대비
exVol = objStockMst.GetHeaderValue(57) #예상체결수량
 
# Send a message to #general channel
slack.chat.post_message('#stock_daily', "종목명 : " + str(name))
slack.chat.post_message('#stock_daily', "종목코드 : " + str(code))
slack.chat.post_message('#stock_daily', "시간 : " + str(time))
slack.chat.post_message('#stock_daily', "종가 : " + str(cprice))
slack.chat.post_message('#stock_daily', "대비 : " + str(diff))
slack.chat.post_message('#stock_daily', "시가 : " + str(open))
slack.chat.post_message('#stock_daily', "고가 : " + str(high))
slack.chat.post_message('#stock_daily', "저가 : " + str(low))
slack.chat.post_message('#stock_daily', "매도호가 : " + str(offer))
slack.chat.post_message('#stock_daily', "매수호가 : " + str(bid))
slack.chat.post_message('#stock_daily', "거래량 : " + str(vol))
slack.chat.post_message('#stock_daily', "거래대금 : " + str(vol_value))
slack.chat.post_message('#stock_daily', '매도호가 : ' + str(offer))
slack.chat.post_message('#stock_daily', '--------------------------------')

if (exFlag == ord('0')):
    slack.chat.post_message('#stock_daily', "장 구분값: 동시호가와 장중 이외의 시간")
elif (exFlag == ord('1')) :
    slack.chat.post_message('#stock_daily', "장 구분값: 동시호가 시간")
elif (exFlag == ord('2')):
    slack.chat.post_message('#stock_daily', "장 구분값: 장중 또는 장종료")

# Send a message to #general channel
slack.chat.post_message('#stock_daily', "예상체결가 대비 수량 : " + exFlag)
slack.chat.post_message('#stock_daily', "예상체결가 : " + exPrice)
slack.chat.post_message('#stock_daily', "예상체결가 대비 : " + exDiff)
slack.chat.post_message('#stock_daily', "예상체결수량 : " + exVol)
slack.chat.post_message('#stock_daily', '===============================')
