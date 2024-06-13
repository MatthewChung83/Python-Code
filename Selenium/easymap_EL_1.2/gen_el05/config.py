import argparse
import datetime
parser = argparse.ArgumentParser(description='外部參數')
parser.add_argument('entitytype', type=str , help='名單名稱')
args = parser.parse_args()
entitytype = args.entitytype

db = {
    'server': 'vnt07.ucs.com',
    'database': 'uis',
    'username': 'pyuser',
    'password': 'Ucredit7607',    
    'MGtb': 'EL_Entity_Managerment',
    'servicetb': 'EL_Service_Datetime',    
    'fromtb': f'EL04',
    'totb': f'EL05',
    'entitytype': entitytype,
    'EntityPhase': 'el05_main.py',
    'EntityPhase_last': 'el04_main.py',
    'EntityPath': ' C:\Py_Project\project\easymap_EL_1.2\gen_el05\\',
    'EntityPath_last': ' C:\Py_Project\project\easymap_EL_1.2\gen_el04\\',
}

wbinfo = {
    'acc_name':'89850389',
    'acc_pwd':'bvi0002',
    #'acc_name':'89603265',
    #'acc_pwd':'vhNzmvnT',
    'url':'https://ep.land.nat.gov.tw/',
    'Drvfile':rf'C:\Py_Project\env\chromedriver_win32\chromedriver',
    'purpose':'EL',
    'disk': 'C:\Py_Project\filesave',
}

APinfo = {
    'Apurl':'https://ocr.ap-mic.com/ocr',
    'imgf':'captcha.jpg',
    'imgp':r'C:\Py_Project\project\easymap_EL_1.2\gen_el04\captcha\captcha.jpg',
}