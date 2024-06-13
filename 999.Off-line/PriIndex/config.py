import datetime
datetime_dt = datetime.datetime.today()# 獲得當地時間
datetime_str = datetime_dt.strftime("%Y/%m/%d")  # 格式化日期


# 正式資料庫資訊
db = {
    'server': 'vnt07',
    'database': 'uis',
    'username': 'pyuser',
    'password': 'Ucredit7607',
    
    # 不動產經紀公司查詢主檔
    'totb1':'PriIndex',
    
    # 名單編號
    'entitytype':'UCS'+'_'+datetime_str.replace('/',''),
}

# 呼叫之API
wbinfo = {
    'url':'https://resim.land.moi.gov.tw/PriQuery/iamqry_234a',
}