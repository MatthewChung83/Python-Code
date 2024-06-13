# 測試資料庫資訊
db = {
    'server': 'vnt07.ucs.com',
    'database': 'uis',
    'username': 'pyuser',
    'password': 'Ucredit7607',

# 裁判書及簡易案件存放table
    'totb1':'FJUD_Default_AD_detail',
    'totb2':'FJUD_Default_AD',
}

# chrome driver 存放本機路徑
wbinfo = {
    'Drvfile':rf'C:\Py_Project\env\chromedriver_win32\chromedriver',}

# 兩種查詢分類
urls = [
    ['https://law.judicial.gov.tw/LAW_MOBILE/default.aspx','簡易案件查詢'],
    ['https://law.judicial.gov.tw/LAW_MOBILE/default.aspx','裁判書查詢'],    
]

# 驗證碼截圖存放路徑
captchainfo = {
    'print_path':r'C:\Py_Project\FJUD_Default_AD\captcha\print_screen.png',
    'captcha_path':rf'C:\Py_Project\FJUD_Default_AD\captcha\captcha.png',
    }

# 查詢起訖日
import datetime
str_date = (datetime.datetime.now() - datetime.timedelta(days = 2)).strftime("%Y/%m/%d")
end_date = (datetime.datetime.now() - datetime.timedelta(days = 0)).strftime("%Y/%m/%d")

# 7種關鍵字查詢
qtypes = ['拍賣','代位','分割','抵押權','變價','塗銷','撤銷']