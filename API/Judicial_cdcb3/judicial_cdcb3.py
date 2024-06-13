# -*- coding: utf-8 -*-
"""
Created on Wed Jan 19 15:08:27 2022

@author: admin
"""


import datetime 
import requests
import json
import re
import pyodbc
import os 
import shutil
import time
import sys

from pathlib import Path
from bs4 import BeautifulSoup
from urllib.parse import urlencode

from config import *
from etl_func import *
from dict import *


server,database,username,password,fromtb,totb = db['server'],db['database'],db['username'],db['password'],db['fromtb'],db['totb']
url,url1= wbinfo['url'],wbinfo['url1']

obs = src_obs(server,username,password,database,fromtb,totb)
mail(obs)

for i in foo(-1,obs-1):
    d=[]
    data = []
    today = str(datetime.datetime.now())[0:-3]
    src = dbfrom(server,username,password,database,fromtb,totb)[0]
    ID = src[1]
    print(ID)
    name = src[2]
    print(name)
    rowid = str(src[9]).replace('None','')
    from requests import Session
    #from fake_useragent import UserAgent
    from bs4 import BeautifulSoup
    
    #uar = UserAgent().random
    req_session = Session()
    #resp = req_session.post(captchacheck_url)
    #cookie = '; '.join( [ '='.join(i) for i in resp.cookies.items()] )
    
    headers={
        #'Cookie': cookie,
        'Host': 'cdcb3.judicial.gov.tw',
        'Origin': 'https://cdcb3.judicial.gov.tw',
        'Referer': 'https://cdcb3.judicial.gov.tw/judbp/wkw/WHD9A01/V2.htm',
        #'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36',
    }
    
    set_token_url = 'https://cdcb3.judicial.gov.tw/judbp/wkw/WHD9A01/V2.htm'
    token_resp = req_session.post(set_token_url, headers=headers)
    
    token = BeautifulSoup( token_resp.text, 'lxml').select('input[name=token]')[0]['value']
    print(token)
    data = {
    'pageNum':'1',
    'pageSize':'20',
    'crtid':'',
    'queryType':'1',
    'clnm':'',
    'clnm_roma':'',
    'idno':ID,
    'sddt_s':'',
    'sddt_e':'',
    'token':token,
    'condition':'undefined'
        }

    #session = requests.session()
    resp = req_session.post(url,headers = headers,data=data)
    soup = BeautifulSoup(resp.text,"lxml")
    print(soup.text)
    d=json.loads(soup.text)
   
    if len(d['data']['dataList']) == 0:
        crtid = ''
        sys = ''
        crmyy = ''
        crmid = ''
        crmno = ''
        crtname = ''
        durdt = ''
        durnm = ''
        filenm = ''
        crm_text = ''
        owner = ''
        attachment_rmk = ''
        attachment_atfilenm = ''
        attachmentnm = ''
        Basis =''
        update_date = today
        note = 'N'
        filename = ''
        docs = (ID,name,crtid,sys,crmyy,crmid,crmno,crtname,durdt,durnm,filenm,crm_text,owner,attachment_rmk,attachment_atfilenm,attachmentnm,Basis,update_date,note,filename)
        judicial_result = judicial(docs)
        if len(rowid) == 0 :
            
            toSQL(judicial_result, totb, server, database, username, password)
            #計算查找筆數，今日調閱大於五千筆則停止     
            exit_o = exit_obs(server,username,password,database,totb)
            if exit_o >= 5000:
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
            delete(server,username,password,database,totb,note,ID,rowid)
            toSQL(judicial_result, totb, server, database, username, password)
            #計算查找筆數，今日調閱大於五千筆則停止     
            exit_o = exit_obs(server,username,password,database,totb)
            if exit_o >= 5000:
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
        d=json.loads(soup.text)
        f=d['data']['dataList']
        for i in range(len(f)):
            crtid=str(f[i]['crtid']).replace('None','')
            sys = str(f[i]['sys']).replace('None','')
            crmyy = str(f[i]['crmyy']).replace('None','')
            crmid = str(f[i]['crmid']).replace('None','')
            crmno = str(f[i]['crmno']).replace('None','')
            crtname = str(f[i]['crtname']).replace('None','')
            durdt = str(f[i]['durdt']).replace('None','')
            durnm = str(f[i]['durnm']).replace('None','')
            filenm = str(f[i]['filenm']).replace('None','')
            crm_text = str(f[i]['crm_text']).replace('None','')
            owner = str(f[i]['owner']).replace('None','')
            try :
                attachment_rmk = str(f[i]['attachment'][0]['rmk']).replace('None','')
                attachment_atfilenm = str(f[i]['attachment'][0]['atfilenm']).replace('None','')
                attachmentnm = str(f[i]['attachmentnm']).replace('None','')
                filename = 'C:\Py_Project\project\judicial_cdcb3\save_file\\'+ID+'_'+crm_text+'_'+str(i)+'.pdf'
                filename1 = Path(filename)
                url3 = attachment_atfilenm
                response = requests.get(url3)
                filename1.write_bytes(response.content)
            except:
                attachment_rmk=''
                attachment_atfilenm=''
                attachmentnm =''
                filename =''
                pass
            data1 = {
                'crtid':crtid,
                'filenm':filenm,
                'condition':'法院別: 全部法院, 公告類型: 消債事件公告, 身分證字號:'+ID,
                'isDialog':'Y',
                }
            
            resp1 = req_session.post(url1,data=data1)
            soup2 = BeautifulSoup(resp1.text.replace('</br>','').replace('<br/>',''),"xml")
            soup1 = soup2.findAll('td')
            if len(soup1[3].text.replace('\u3000','').replace('\n','').replace('\t','').replace(' ','')) > 4000:
                Basis = ''
            else:
                Basis = soup1[3].text.replace('\u3000','').replace('\n','').replace('\t','').replace(' ','')
            Basis = ''
            update_date = today
            note = 'Y'

            docs = (ID,name,crtid,sys,crmyy,crmid,crmno,crtname,durdt,durnm,filenm,crm_text,owner,attachment_rmk,attachment_atfilenm,attachmentnm,Basis,update_date,note,filename)
            judicial_result = judicial(docs)
            if len(rowid) == 0 :
            
                toSQL(judicial_result, totb, server, database, username, password)
                
                #計算查找筆數，今日調閱大於五千筆則停止     
                exit_o = exit_obs(server,username,password,database,totb)
                if exit_o >= 5000:
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
                delete(server,username,password,database,totb,note,ID,rowid)
                toSQL(judicial_result, totb, server, database, username, password)
                
                #計算查找筆數，今日調閱大於五千筆則停止     
                exit_o = exit_obs(server,username,password,database,totb)
                if exit_o >= 5000:
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
                    
  

    
if obs > 0 :
    errormail(obs)

elif obs ==0 :
    errormail(obs)
try:
 
    file_source = rf'C:\Py_Project\project\judicial_cdcb3\save_file\\'
    file_destination = rf'\\fortune\Cashfile\UCS\judicial_cdcb3\\'
     
    get_files = os.listdir(file_source)
     
    for g in get_files:
        shutil.copy(file_source + g, file_destination)
   
    
except:
    #ERR_mail(obs)
    pass
