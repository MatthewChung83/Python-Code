from operator import itemgetter, attrgetter
from datetime import datetime, timedelta
import datetime
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
import pandas as pd
import requests
import time
import sys


from etl_func import *
from config import *
from etl_func import *
from dict import *

server,database,username,password,totb1,totb2 = db['server'],db['database'],db['username'],db['password'],db['totb1'],db['totb2']
url = wbinfo['url']

check = check(server, username, password, database)
obs =  src_obs(server,username,password,database,totb1,totb2,check)

mail(obs)
print(obs)
print(check)
for i in foo(-1,obs-1):
    
    src = dbfrom(server,username,password,database,totb1,totb2,check)[0]
    #  id = 'A120373390'
    id = src[1]
    print(id)
    rowid = str(src[9]).replace('None','')
    print(rowid)
    today = str(datetime.datetime.now())[0:-3]
    soup = parse(url,id)

    response,status = [],'N'
    if soup.find(class_ = 'Standard'):
        
        status = 'Y'
        for t in soup.find(class_ = 'Standard').find_all('tr'):
            if t.find('th') and t.find('td'):
                response.append(t.find('th').text + t.find('td').text)        
    
    response = '&&'.join(response)
    print(response)
    name = src[2]
    update_date = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    docs = (name,id,response,update_date,status)
    doc = indata(docs)
    
    if len(rowid) >0 :
        update(server,username,password,database,totb2,status,id,today,rowid,response)
    else:
        toSQL(doc, totb2, server, database, username, password)

    #計算查找筆數，今日調閱大於五千筆則停止     
    exit_o = exit_obs(server,username,password,database,totb2)
    if exit_o >= 6000:
        sys.exit()
    else:
        pass
    
    #計算查找時間，因OCR作業於下午1點後，查找至中午12點停止
    date_time=time.localtime()
    now=time.strftime("%Y-%m-%d %H:%M:%S",date_time)
    batstoptime = time.strftime("%Y-%m-%d",date_time)+' 11:59:59'
    if now <= batstoptime:
        print('sys exec ; now=',now,' batstoptime=',batstoptime)
        pass
    else:
        print('sys stop ; now=',now,' batstoptime=',batstoptime)
        sys.exit()




    
obs = src_obs(server,username,password,database,totb1,totb2,check)
errormail(obs)
if obs > 0 :
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update UCS_Source
    set F_OBS = '{obs}'
    where Engine = 'Insurance_UnitQuery' and type = '{check}'
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
elif obs ==0 :
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update UCS_Source
    set status = 'Done',F_OBS = '{obs}'
    where Engine = 'Insurance_UnitQuery' and type = '{check}'
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()