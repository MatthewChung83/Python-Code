# -*- coding: utf-8 -*-
db = {
    # 資料回寫主機之資訊
    'server': '10.10.0.94',
    'database': 'CL_Daily',
    'username': 'CLUSER',
    'password': 'Ucredit7607',
    'fromtb': 'base_case',
    # 監理所駕照查詢主檔
    'totb':'CC95T',
}

wbinfo = {
    # post url
    'url':'https://eservice.wdasec.gov.tw/WWW/CC95T',

    # captcha url
    'captchaImg':'https://eservice.wdasec.gov.tw/Login/GetValidateCode',
}

pics = {
    # captcha path
    'imgp' : r'C:\Py_Project\project\CC95T\ValidateCode.png'
}