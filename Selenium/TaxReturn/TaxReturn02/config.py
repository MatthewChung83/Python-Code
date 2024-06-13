import argparse

parser = argparse.ArgumentParser(description='外部參數')
parser.add_argument('entitytype', type=str , help='名單名稱')      # 名單名稱
args = parser.parse_args()
entitytype = args.entitytype

db = {
    'server': '10.10.0.94',
    'database': 'CL_Daily',
    'username': 'CLUSER',
    'password': 'Ucredit7607',
    'entitytype': entitytype,
    'fromtb':'taxreturntb',
    'totb':'taxreturntb',
}

APinfo = {
    'Apurl':'https://ocr.ap-mic.com/ocr',
    'imgf1':'captcha1.jpg',
    'imgf2':'captcha2.jpg',
    'imgp1':r'C:\Py_Project\project\TaxReturn02\picture\captcha1.jpg',
    'imgp2':r'C:\Py_Project\project\TaxReturn02\picture\captcha2.jpg',
}