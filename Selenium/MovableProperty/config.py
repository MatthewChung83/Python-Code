import argparse
parser = argparse.ArgumentParser(description='外部參數')
parser.add_argument('entitytype', type=str , help='名單名稱')      # 名單名稱
args = parser.parse_args()
entitytype = args.entitytype

db = {
    #'server': '10.90.0.62',
    #'server': 'vnt07',
    #'database': 'uis',
    #'username': 'pyuser',
    'server': '10.10.0.94',
    'database': 'CL_Daily',
    'username': 'CLUSER',
    'password': 'Ucredit7607',      
    'fromtb': 'PropertySecuredtb',
    'totb': 'PropertySecuredtbinfo',
}

wbinfo = {
    'url':'https://ppstrq.nat.gov.tw/pps/pubQuery/PropertyQuery/propertyQuery.do',
    'Drvfile':rf'C:\Py_Project\env\chromedriver_win32\chromedriver.exe',
    'disk': rf'C:\Py_Project\project\MovableProperty\file',
}