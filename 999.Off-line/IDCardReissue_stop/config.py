import datetime
datetime_dt = datetime.datetime.today()# 獲得當地時間
datetime_str = datetime_dt.strftime("%Y/%m/%d")  # 格式化日期
#print(datetime_str.replace('/',''))
db = {
    # 資料回寫主機之資訊
    'server': 'vnt07',
    'database': 'uis',
    'username': 'pyuser',
    'password': 'Ucredit7607',
    
    # 身分證查詢主檔
    'totb1':'IDCardReissue',
    
    # 名單編號
    #'entitytype':'UCS'+'_'+datetime_str.replace('/',''),
    'entitytype':'matthew_test',
}

#辨識碼儲存位置
APinfo = {
    'imgf':'captcha.jpg',
    'imgp':r'C:\Users\admin\Desktop\test\captcha.jpg',
        }

#查找網址&chrome執行檔
wbinfo = {   
    'url':'https://www.ris.gov.tw/apply-idCard/app/idcard/IDCardReissue/main',
    'Drvfile':rf'C:\Py_Project\env\chromedriver_win32\chromedriver',
        }