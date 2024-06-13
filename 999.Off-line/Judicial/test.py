from operator import itemgetter, attrgetter
from datetime import datetime, timedelta
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
import pandas as pd
import requests
import time

from etl_func import *
from config import *
from etl_func import *
from dict import *

server,database,username,password,totb1,totb2 = db['server'],db['database'],db['username'],db['password'],db['totb1'],db['totb2']
url = wbinfo['url']

# 分關鍵字查詢
for i in range(len(qtypes)):
    qtype = qtypes[i]
    
    data = {
        'Action': 'Qeury',
        'lstoken': '9996409954',
        'Q_DMDeptMainID':'',
        'TBOXDMPostDateS': str_date,
        'TBOXDMPostDateE': end_date,
        'Q_DTACaseNumber':'',
        'Q_DMTitle': '',
        'Q_DTACatCode1':'',
        'Q_DTAColumn4':'',
        'Q_DMCatCode':'',
        'Q_DMBody': qtype,
        'BtnSubmit': '查詢',
        }
    
    resp = requests.post(url,data=data)
    soup = BeautifulSoup(resp.text,"lxml")
    
    print(soup)