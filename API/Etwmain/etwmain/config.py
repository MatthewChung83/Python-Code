# -*- coding: utf-8 -*-
"""
Created on Mon May  9 18:34:52 2022

@author: admin
"""


db = {
    'server': 'RICH',
    'database': 'UCS_ReportDB',
    'username': 'FRUSER',
    'password': '1qaz@WSX',
    'fromtb':'UCS_ReportDB.dbo.INS_LegalG',
    'fromtb1':'UCS_ReportDB.dbo.INS_LegalG_Sub',
    'totb':'treasure.skiptrace.dbo.Legal_G_List',
    
}

wbinfo = {
    'url':'https://www.etax.nat.gov.tw/etwmain/etw113w1/name/query',
    'Drvfile':rf'C:\Py_Project\env\chromedriver_win32\chromedriver',
    'imgp':rf'C:\Py_Project\project\etwmain\etwmain\captcha.jpg',
}
