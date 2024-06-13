import argparse
import datetime
parser = argparse.ArgumentParser(description='外部參數')
parser.add_argument('entitytype', type=str , help='名單名稱')      # 名單名稱
parser.add_argument('batstoptime', type=str , help='批次停止時間')  # 批次停止時間
args = parser.parse_args()
entitytype = args.entitytype
batstoptime = args.batstoptime

db = {
    'server': 'vnt07.ucs.com',
    'database': 'uis',
    'username': 'pyuser',
    'password': 'Ucredit7607',
    'fromtb':'EL02',
    'totb': 'EL03',
    'entitytype': entitytype,
    'batstoptime': batstoptime,
    'todtltb': 'EL03DTL',
    'EntityPhase':'el03_main.py',
    'EntityPhase_next':'el03_main.py',
    'EntityPath_next':' C:\Py_Project\project\easymap_EL_1.2\gen_el03\\',
    'MGtb':'EL_Entity_Managerment',
}