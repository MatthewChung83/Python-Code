import datetime
datetime_dt = datetime.datetime.today()# 獲得當地時間
datetime_str = datetime_dt.strftime("%Y/%m/%d")  # 格式化日期




db = {
    'server': 'vnt07',
    'database': 'uis',
    'username': 'pyuser',
    'password': 'Ucredit7607',
    'totb':'tmnewa',
    'fromtb':'base_case',
}

wbinfo = {
    'url':'https://www.tmnewa.com.tw/B2C_V2/commonweb/covid19-insurance-service.aspx?fk=45e3015a5a8a495e9385e816f71c3c4b',
    'Drvfile':rf'C:\Py_Project\env\chromedriver_win32\chromedriver',
    'imgp':r'C:\Py_Project\project\tmnewa\captcha.jpg',
}