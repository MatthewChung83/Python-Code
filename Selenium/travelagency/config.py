import datetime
datetime_dt = datetime.datetime.today()# 獲得當地時間
datetime_str = datetime_dt.strftime("%Y/%m/%d")  # 格式化日期




db = {
    'server': 'vnt07',
    'database': 'uis',
    'username': 'pyuser',
    'password': 'Ucredit7607',
    'totb':'travelagency',
    'fromtb':'base_case',
}

wbinfo = {
    'url':'https://travelagency.tbroc.gov.tw/DataQuery/Emp_JobNow.aspx',
    'Drvfile':rf'C:\Py_Project\env\chromedriver_win32\chromedriver',
    'imgp':r'C:\Py_Project\project\tmnewa\captcha.jpg',
}