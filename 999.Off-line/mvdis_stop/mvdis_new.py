# -*- coding: utf-8 -*-
"""
Created on Tue Feb 15 14:07:33 2022

@author: admin
"""
from bs4 import BeautifulSoup
import requests
import ddddocr
import time
import re
import datetime

from etl_func import *
from config import *
from dict import *



server,database,username,password,fromtb,totb = db['server'],db['database'],db['username'],db['password'],db['fromtb'],db['totb']
url,captchaImg,imgp= wbinfo['url'],wbinfo['captchaImg'],wbinfo['imgp']
check = check(server, username, password, database)

obs = src_obs(server,username,password,database,fromtb,totb,check)
print(obs)
mail(obs)
q=0
for i in foo(-1,obs-1):

    obs = src_obs(server,username,password,database,fromtb,totb,check)
    src = dbfrom(server,username,password,database,fromtb,totb,check)[0]
    num(obs,server,username,password,database,fromtb,totb,check)
    today = str(datetime.datetime.now())[0:-3]
    #rowid = i
    Name = src[2]
    ID = src[1]
    birthday = src[6]
    a=birthday.split('-',1)
    rowid = str(src[0]).replace('None','')
    if len(ID) == 10:
        if len(a[0]) == 3:
            bir = birthday.replace('/','').replace('-','')
        else : 
            bir = '0'+birthday.replace('/','').replace('-','')
    else:
        note = 'N'
        update_date = today
        docs = (Name,ID,birthday,update_date,note)
        dead_result = dead(docs)
        if len(rowid) > 0:
            update(server,username,password,database,totb,note,ID,today,rowid)
        else:
            toSQL(dead_result, totb, server, database, username, password)
        continue
    print(ID)
    print(bir)
    resp = capcha_resp(url,captchaImg,imgp,q,ID,bir)
    msg_box_warning = resp.find(class_ = 'msg_box_warning').text
    print(msg_box_warning)
    
    if '請確認您輸入的證號及生日是否正確。' in msg_box_warning:
        note = 'N'
        update_date = today
        docs = (Name,ID,birthday,update_date,note)
        dead_result = dead(docs)
        if len(rowid) > 0:
            update(server,username,password,database,totb,note,ID,today,rowid)
        else:
            toSQL(dead_result, totb, server, database, username, password)
        continue
    else:
        pass
    driver_info = resp.find(class_ = 'tb_list_std')
    try:
        driver_status = resp.find('tbody').text
    except:
        note = 'N'
        update_date = today
        docs = (Name,ID,birthday,update_date,note)
        dead_result = dead(docs)
        if len(rowid) > 0:
            update(server,username,password,database,totb,note,ID,today,rowid)
        else:
            toSQL(dead_result, totb, server, database, username, password)
        continue
    #print(driver_info)
    #print(driver_status)
    if '免換照' in driver_status :
        note = 'N'
        update_date = today
        docs = (Name,ID,birthday,update_date,note)
        dead_result = dead(docs)
        if len(rowid) > 0:
            update(server,username,password,database,totb,note,ID,today,rowid)
        else:
            toSQL(dead_result, totb, server, database, username, password)
    elif '此車非活車' in driver_status:
        note = 'Y'
        update_date = today
        docs = (Name,ID,birthday,update_date,note)
        dead_result = dead(docs)
        if len(rowid) > 0:
            update(server,username,password,database,totb,note,ID,today,rowid)
        else:
            toSQL(dead_result, totb, server, database, username, password)
    else  :
        note = 'N'
        update_date = today
        docs = (Name,ID,birthday,update_date,note)
        dead_result = dead(docs)
        if len(rowid) > 0:
            update(server,username,password,database,totb,note,ID,today,rowid)
        else:
            toSQL(dead_result, totb, server, database, username, password)

    
obs = src_obs(server,username,password,database,fromtb,totb,check)
if obs > 0 :
    
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update UCS_Source
    set F_OBS = '{obs}'
    where Engine = 'dead' and type = '{check}'
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    errormail(obs)
    
elif obs ==0 :
    
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update UCS_Source
    set status = 'Done'
    where Engine = 'dead' and type = '{check}'
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()    
    errormail(obs)
    driver.quit()