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
import pdfkit
from bs4 import BeautifulSoup
#from pyquery import PyQuery as pq
from urllib.parse import urlencode
import urllib

from config import *
from etl_func import *
from dict import *


server,database,username,password,fromtb,totb = db['server'],db['database'],db['username'],db['password'],db['fromtb'],db['totb']
url,url1= wbinfo['url'],wbinfo['url1']
check = check(server, username, password, database)
print(check)
obs = src_obs(server,username,password,database,fromtb,totb,check)
mail(obs)

for i in foo(-1,obs-1):
    today = str(datetime.datetime.now())[0:-3]
    src = dbfrom(server,username,password,database,fromtb,totb,check)[0]
    ID = src[1]
    print(ID)
    name = src[0]
    print(name)
    rowid = ''
    print(rowid)
    data = {
    'crtid':'',
    'kdid':'',
    'clnm':name,
    'clnm_roma':'',
    'idno':ID,
    'sddt_s':'',
    'sddt_e':'',
    'condition':'法院別: 全部法院, 類別: 全部, 當事人: '+name+', 身分證字號: '+ID
    }
    session = requests.session()
    resp = session.post(url,data=data)
    soup = BeautifulSoup(resp.text,"lxml")
    d=json.loads(soup.text)
    if len(d['data']['dataList']) == 0:
        announcement = ''
        post_date = ''
        register_no = ''
        keynote = ''
        Basis = ''
        Matters = ''
        update_date = today
        note = 'N'
        flag = '查無資料'
        file = ''
        docs = (name,ID,announcement,post_date,register_no,keynote,Basis,Matters,update_date,note,flag,file)
        print(judicial(docs))
        judicial_result = judicial(docs)
        print(judicial_result)
        toSQL(judicial_result, totb, server, database, username, password)
        continue
    else:
        
        d=json.loads(soup.text)
        f=d['data']['dataList']
        print(f)
        print(range(len(f)))
        for i in range(len(f)):
            crtid=f[i]['crtid']
            print(crtid)
            filenm=f[i]['filenm']
            print(filenm)
            durcd=f[i]['durcd']
            print(durcd)
            seqno=f[i]['seqno']
            item=f[i]['item']
            data1 = {
                'crtid':crtid,
                'filenm':filenm,
                'durcd':durcd,
                'condition':'法院別: 全部法院, 類別: 全部, 當事人: '+', 身分證字號: '+ID,
                'seqno':seqno,
                'isDialog':'Y',
                }
            print(data1)
            condition =urllib.parse.quote('法院別: 全部法院, 類別: 全部, 當事人: '+', 身分證字號: '+ID).replace('%20','+')
            Basis='https://domestic.judicial.gov.tw/judbp/wkw/WHD9HN01/VIEW.htm?crtid='+crtid+'&filenm='+filenm.replace('/','%2F')+'&durcd='+durcd+'&condition='+condition+'&seqno='+seqno+'&isDialog=Y'
            config = pdfkit.configuration(wkhtmltopdf="C:\Program Files\wkhtmltopdf\wkhtmltopdf.exe")
            file =  rf'C:\Py_Project\project\judicial_fam_TSTAR\20220826\\'+name+'_'+ID+'_'+str(i)+'.pdf'
            pdfkit.from_url(Basis,file,configuration=config)
            resp = session.post(url1,data=data1)
            soup = BeautifulSoup(resp.text.replace('</br>','').replace('<br/>',''),"xml")
            soup1 = soup.findAll('td')
            print(soup1[3].text)
            text = soup1[3].text.replace('			例','例')
            if '遺產管理人' in item :
                flag = item
                note = 'Y'
            elif '遺產清冊' in item:
                flag = item
                note = 'Y'
            elif '拋棄繼承' in item:
                flag = item
                note = 'Y'
            else :
                flag = item
                note = 'N'
            try:
                announcement = re.search( r'例稿名稱：(.*)',text , re.M|re.I).group()
            except:
                announcement = ''
                pass
            try:
                post_date = re.search( r'發文日期：(.*)',text , re.M|re.I).group()
            except:
                post_date = ''
                pass
            try:
                register_no = re.search( r'發文字號：(.*)',text , re.M|re.I).group()
            except:
                register_no = ''
                pass
            try:
                keynote = re.search( r'主[\s]*旨：(.*)',text , re.M|re.I).group()
            except:
                keynote = ''
                pass
            #file = ''
            Matters = text.replace('\u3000','').replace('┌','').replace('─','').replace('┬','').replace('│','').replace('┼','').replace('├','').replace('┤','').replace('─','').replace('└','').replace('┴','').replace('┘','')
            update_date = today
            #note = 'Y'
            #flag = '查無資料'
            docs = (name,ID,announcement,post_date,register_no,keynote,Basis,Matters,update_date,note,flag,file)
            judicial_result = judicial(docs)
            print(judicial_result)
            toSQL(judicial_result, totb, server, database, username, password)
            continue
    
if obs > 0 :
    errormail(obs)
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update UCS_Source
    set F_OBS = '{obs}'
    where Engine = 'judicial_fam' and type = '{check}'
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
elif obs ==0 :
    errormail(obs)
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    update UCS_Source
    set status = 'Done',F_OBS = '{obs}'
    where Engine = 'judicial_fam' and type = '{check}'
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()