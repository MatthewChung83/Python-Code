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
    'fromtb':'el00',
    'totb': 'el02',
    'entitytype': entitytype,
    'batstoptime': batstoptime,
}