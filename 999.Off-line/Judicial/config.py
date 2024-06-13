from datetime import datetime, timedelta
import datetime
# 指定查詢公告日期，-0 為當下日期
#str_date = f"{datetime.now().year - 1911}/{datetime.now().strftime('%m')-3}/{datetime.now().day - 0}"
#end_date = f"{datetime.now().year - 1911}/{datetime.now().strftime('%m')}/{datetime.now().day - 0}"
str_date = (datetime.datetime.now() - datetime.timedelta(days = 92)).strftime("%Y/%m/%d")
end_date = (datetime.datetime.now() - datetime.timedelta(days = 0)).strftime("%Y/%m/%d")

# 測試資料庫資訊
db = {
    'server': 'vnt07.ucs.com',
    'database': 'uis',
    'username': 'pyuser',
    'password': 'Ucredit7607',
    
    # 法文主檔
    'totb1':'judicial',
    
    # 債務人檔
    'totb2':'judicial_person',
}

# 呼叫之API
wbinfo = {
    'url':'https://www.judicial.gov.tw/tw/lp-139-1.html',
}