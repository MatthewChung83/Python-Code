# -*- coding: utf-8 -*-
"""
Created on Thu Jan 20 22:42:02 2022

@author: admin
"""


import datetime
import time
import requests
import sys

from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from requests import Session



from config import *
from etl_func import *
from dict import *

server,database,username,password,fromtb,totb,checktb= db['server'],db['database'],db['username'],db['password'],db['fromtb'],db['totb'],db['checktb']
url = wbinfo['url']


obs = src_obs(server,username,password,database,fromtb,totb)

print(src_obs(server, username, password, database,fromtb,totb))
mail(obs)
for i in foo(-1,obs-1):
    src = dbfrom(server,username,password,database,fromtb,totb)[0]
    #rowid = i
    Name = src[2]
    ID = src[1]
    rowid = str(src[9]).replace('None','')
    print(rowid)
    print(ID)
    data = {
        'mall_idn': ID, 
        'cer_type': '0',
        }
    r = requests.post(url=url, data = data)

    soup = BeautifulSoup(r.text,"lxml")
    
    name = soup.findAll('head')
    #print(name)
    obs = src_obs(server,username,password,database,fromtb,totb)
    try:
        aa=name[0].text.strip()
        print(aa)
    
        today = str(datetime.datetime.now())[0:-3]

        if '查無資料' in aa :
            note = 'N'
            member_type=''
            register_no=''
            Receipt=''
            Certificate=''
            update_date = today
            docs = (Name,ID,member_type,register_no,Receipt,Certificate,update_date,note)
            resim_result = resim(docs)
            if len(rowid) > 0:
                update(server,username,password,database,totb,note,ID,today,rowid,member_type,register_no,Receipt,Certificate)
                #計算查找筆數，今日調閱大於五千筆則停止     
                exit_o = exit_obs(server,username,password,database,totb)
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
            else:
                toSQL(resim_result, totb, server, database, username, password)
                #計算查找筆數，今日調閱大於五千筆則停止     
                exit_o = exit_obs(server,username,password,database,totb)
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
            #print(src_obs(server, username, password, database))
            continue
    except :
    #else:   
        soup = BeautifulSoup(r.text,"lxml")
        name = soup.findAll('tr')
        a= name[2].text.strip().split('\n')
        member_type =a[0]
        Name = a[1]
        Receipt=a[2]
        register_no = a[3]
        Certificate = a[4]
        if '證照種類' in member_type :
            note = 'N'
            continue
        else:
            note = 'Y'
            update_date = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            docs = (Name,ID,member_type,register_no,Receipt,Certificate,update_date,note)
            resim_result = resim(docs)
            #print(resim_result)
            print(a)
            if len(rowid) > 0:
                update(server,username,password,database,totb,note,ID,update_date,rowid,member_type,register_no,Receipt,Certificate)
                #計算查找筆數，今日調閱大於五千筆則停止     
                exit_o = exit_obs(server,username,password,database,totb)
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
            else:
                toSQL(resim_result, totb, server, database, username, password)
                #計算查找筆數，今日調閱大於五千筆則停止     
                exit_o = exit_obs(server,username,password,database,totb)
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
            #print(src_obs(server, username, password, database))
            pass





        


