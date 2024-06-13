db = {
    # 資料回寫主機之資訊
    'server': 'vnt07',
    'database': 'uis',
    'username': 'pyuser',
    'password': 'Ucredit7607',
    
    # 監理所駕照查詢主檔
    'totb1':'LicensePenalty_once',
    
    # 名單編號
    'entitytype':'Persontb',
}

wbinfo = {
    # post url
    'url':'https://www.mvdis.gov.tw/m3-emv-vil/vil/driverLicensePenalty',

    # captcha url
    'captchaImg':'https://www.mvdis.gov.tw/m3-emv-vil/captchaImg.jpg',
}

pics = {
    # captcha path
    'imgp' : r'C:\Py_Project\project\LicensePenalty_test\valcode.png'
}