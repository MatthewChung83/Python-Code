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

server,database,username,password,fromtb,totb= db['server'],db['database'],db['username'],db['password'],db['fromtb'],db['totb']
url = wbinfo['url']


obs = src_obs(server,username,password,database,fromtb,totb)

print(src_obs(server, username, password, database,fromtb,totb))
#mail(obs)
for i in foo(-1,obs-1):
    error = ''
    if i ==10000:
        break
    updatetime = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    src = dbfrom(server,username,password,database,fromtb,totb)[0]
    #rowid = i
    Name = src[1]
    ID = src[0]
    rowid = str(src[2]).replace('None','')
    name_web_01 = src[3]
    city = src[4]
    license_num = src[5]
    license_valid = src[6]
    phone = src[7]
    juild = src[8]
    print(rowid)
    print(ID)
    data = {
    'acert_pract_county_cd':'', 
    'acert_pract_area_cd':'', 
    'acert_name': Name,
    'acert_idn': ID,
    'acert_cer_wn_year':'', 
    'acert_cer_wn_word':'', 
    'acert_cer_wn_no':'', 
    'acert_pract_wn_year':'', 
    'acert_pract_wn_word':'', 
    'acert_pract_wn_no':'', 
    'acert_pract_org_name':'' ,
        }
    r = requests.post(url=url, data = data)

    soup = BeautifulSoup(r.text,"lxml")
    obs = src_obs(server,username,password,database,fromtb,totb)
    try:
        error = soup.find('head').text.strip().replace(' ','').split('\n')[0]
    except:
        error = ''
        pass
    if '查無資料' in error:
        note = 'N'
        Name_web = ''
        city_web = ''
        license_num_web = ''
        license_valid_web = ''
        office_name_web = ''
        phone_web = ''
        juild_web = ''
    else:
        note = 'Y'
        Name_web = soup.findAll('tr')[1].text.strip().replace(' ','').split('\n')[4]
        city_web = soup.findAll('tr')[1].text.strip().replace(' ','').split('\n')[5]
        license_num_web = soup.findAll('tr')[1].text.strip().replace(' ','').split('\n')[6]
        license_valid_web = soup.findAll('tr')[1].text.strip().replace(' ','').split('\n')[7]
        office_name_web = soup.findAll('tr')[1].text.strip().replace(' ','').split('\n')[8]
        phone_web = soup.findAll('tr')[1].text.strip().replace(' ','').split('\n')[9]
        juild_web = soup.findAll('tr')[1].text.strip().replace(' ','').split('\n')[10]
    
    print(note)
    
    if note == 'N' and len(rowid)>0 :
        update(server, username, password, database, totb, note, name, rowid, Name_web, city, license_num, license_valid, office_name, phone, juild, ID, today)
        #計算查找筆數，今日調閱大於五千筆則停止     
        exit_o = exit_obs(server,username,password,database,totb)
        if exit_o >= 15000:
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
    elif note == 'Y' and  name_web_01 == Name_web and city == city_web and license_num == license_num_web and license_valid == license_valid_web and office_name == office_name_web and phone == phone_web and juild == juild_web:
        update(server, username, password, database, totb, note, name, rowid, Name_web, city, license_num, license_valid, office_name, phone, juild, ID, today)
        #計算查找筆數，今日調閱大於五千筆則停止     
        exit_o = exit_obs(server,username,password,database,totb)
        if exit_o >= 15000:
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
        docs = (ID,Name,Name_web,city_web,license_num_web,license_valid_web,office_name_web,phone_web,juild_web,updatetime,note)
        resim_AgentIndex_result = resim_AgentIndex(docs)
        toSQL(resim_AgentIndex_result, totb, server, database, username, password)
        #計算查找筆數，今日調閱大於五千筆則停止     
        exit_o = exit_obs(server,username,password,database,totb)
        if exit_o >= 15000:
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



        


