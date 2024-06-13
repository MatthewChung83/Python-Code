import datetime
datetime_dt = datetime.datetime.today()# 獲得當地時間
datetime_str = datetime_dt.strftime("%Y/%m/%d")  # 格式化日期
#print(datetime_str.replace('/',''))
db = {
    # 資料回寫主機之資訊
    'server': '10.10.0.94',
    'database': 'CL_Daily',
    'username': 'CLUSER',
    'password': 'Ucredit7607',
    
    # 監理所駕照查詢主檔
    'totb1':'LicensePenalty',
    
    # 名單編號
    'entitytype':'UCS'+'_'+datetime_str.replace('/',''),
}

wbinfo = {
    # post url
    'url':'https://www.mvdis.gov.tw/m3-emv-vil/vil/driverLicensePenalty',

    # captcha url
    'captchaImg':'https://www.mvdis.gov.tw/m3-emv-vil/captchaImg.jpg',
}

pics = {
    # captcha path
    'imgp' : r'C:\Py_Project\project\LicensePenalty\valcode.png'
}