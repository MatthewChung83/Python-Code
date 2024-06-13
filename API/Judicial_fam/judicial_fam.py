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
import time
import sys
import os

from bs4 import BeautifulSoup
#from pyquery import PyQuery as pq
from urllib.parse import urlencode
import urllib

from config import *
from etl_func import *
from dict import *


server,database,username,password,fromtb,totb = db['server'],db['database'],db['username'],db['password'],db['fromtb'],db['totb']
url,url1= wbinfo['url'],wbinfo['url1']

obs = src_obs(server,username,password,database,fromtb,totb)
print(obs)
mail(obs)
src = dbfrom(server,username,password,database,fromtb,totb)
for i in range(len(src)):
    today = str(datetime.datetime.now())[0:-3]
    #src = dbfrom(server,username,password,database,fromtb,totb)[0]
    ID = src[i][1]
    print(ID)
    name = src[i][2]
    print(name)
    rowid = str(src[i][9]).replace('None','')
    print(rowid)
    from requests import Session
    #from fake_useragent import UserAgent
    from bs4 import BeautifulSoup
    
    #uar = UserAgent().random
    req_session = Session()
    #resp = req_session.post(captchacheck_url)
    #cookie = '; '.join( [ '='.join(i) for i in resp.cookies.items()] )
    
    headers={
        #'Cookie': cookie,
        'Host': 'domestic.judicial.gov.tw',
        'Origin': 'https://domestic.judicial.gov.tw',
        'Referer': 'https://domestic.judicial.gov.tw/judbp/wkw/WHD9HN01/V2.htm',
        #'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36',
    }
    
    set_token_url = 'https://domestic.judicial.gov.tw/judbp/wkw/WHD9HN01/V2.htm'
    token_resp = req_session.post(set_token_url, headers=headers)
    token = BeautifulSoup( token_resp.text, 'lxml').select('input[name=token]')[0]['value']
    print(token)
    
    data = {
        'crtid':'',
        'kdid':'',
        'clnm':name,
        'clnm_roma':'',
        'idno':ID,
        'sddt_s':'',
        'sddt_e':'',
        'token':token,
        'condition':'法院別: 全部法院, 類別: 全部, 身分證字號: '+ID
    }
    
    resp=req_session.post(url, headers=headers, data=data)
    soup = BeautifulSoup(resp.text,"lxml")
    print(soup)
    try:
        d=json.loads(soup.text)
        res = d['data']['dataList']
    except:
        res = ''
        pass
    if len(res) == 0:
        announcement = ''
        post_date = ''
        register_no = ''
        keynote = ''
        Basis = ''
        Matters = ''
        update_date = today
        note = 'N'
        item = ''
        flag = '查無資料'
        note2={"status":"","Death_date":""}
        note2 = json.dumps(note2,ensure_ascii=False)
        docs = (name,ID,announcement,post_date,register_no,keynote,Basis,Matters,update_date,note,flag,note2)
        judicial_result = judicial(docs)
        #print(judicial_result)
        if len(rowid) == 0 :
            
            toSQL(judicial_result, totb, server, database, username, password)
            #計算查找筆數，今日調閱大於五千筆則停止     
            exit_o = exit_obs(server,username,password,database,totb)
            if exit_o >= 100000:
                sys.exit()
            else:
                pass
            
           
        else:
            delete(server,username,password,database,totb,note,ID,rowid,register_no,flag)
            toSQL(judicial_result, totb, server, database, username, password)
            exit_o = exit_obs(server,username,password,database,totb)
            if exit_o >= 100000:
                sys.exit()
            else:
                pass
            
            
            
    else:
        
        d=json.loads(soup.text)
        f=d['data']['dataList']
        #print(f)
        #print(range(len(f)))
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
                'filenm':str(filenm).replace('None',''),
                'durcd':str(durcd).replace('None',''),
                'condition':'法院別: 全部法院, 類別: 全部, 當事人: '+name+', 身分證字號: '+ID,
                'seqno':str(seqno).replace('None',''),
                'isDialog':'Y',
                }
            condition =urllib.parse.quote('法院別: 全部法院, 類別: 全部, 當事人: '+name+', 身分證字號: '+ID).replace('%20','+')
            #print('crtid='+crtid+'&filenm='+filenm.replace('/','%2F')+'&durcd='+durcd+'&condition='+condition+'&seqno='+seqno+'&isDialog=Y')
            #print(data1)
            resp = req_session.post(url1,data=data1)
            soup = BeautifulSoup(resp.text.replace('</br>','').replace('<br/>',''),"xml")
            soup1 = soup.findAll('td')
            try:
                print(soup1[3].text)
                text = soup1[3].text.replace('			例','例')
            except:
                text = ''
                pass
            if '遺產管理人' in item :
                item = '遺產管理人'
                note = 'Y'
            elif '遺產清冊' in item:
                item = '陳報遺產清冊'
                note = 'Y'
            elif '拋棄繼承' in item:
                item = '拋棄繼承'
                note = 'Y'
            else :
                item = item
                note = 'N'
            try:
                announcement = re.search( r'例稿名稱：(.*)',text , re.M|re.I).group().replace(' ','').replace('\n','').replace('─','').replace('┼','').replace('┌','').replace('├','').replace('┬','').replace('┤','').replace('┐','').replace('│','').replace('┤','').replace('┘','').replace('└','').replace('　','')
            except:
                announcement = ''
                pass
            try:
                post_date = re.search( r'發文日期：(.*)',text , re.M|re.I).group().replace(' ','').replace('\n','').replace('─','').replace('┼','').replace('┌','').replace('├','').replace('┬','').replace('┤','').replace('┐','').replace('│','').replace('┤','').replace('┘','').replace('└','').replace('　','')
            except:
                post_date = ''
                pass
            try:
                register_no = re.search( r'發文字號：(.*)',text , re.M|re.I).group().replace(' ','').replace('\n','').replace('─','').replace('┼','').replace('┌','').replace('├','').replace('┬','').replace('┤','').replace('┐','').replace('│','').replace('┤','').replace('┘','').replace('└','').replace('　','')
            except:
                register_no = ''
                pass
            try:
                keynote = re.search( r'主[\s]*旨：(.*)',text , re.M|re.I).group().replace(' ','').replace('\n','').replace('─','').replace('┼','').replace('┌','').replace('├','').replace('┬','').replace('┤','').replace('┐','').replace('│','').replace('┤','').replace('┘','').replace('└','').replace('　','')
            except:
                keynote = ''
                pass
            Basis = ''
            Matters = text.replace('\u3000','').replace('┌','').replace('─','').replace('┬','').replace('│','').replace('┼','').replace('├','').replace('┤','').replace('─','').replace('└','').replace('┴','').replace('┘','').replace(' ','').replace('\n','').replace(' ','').replace('　','')
            update_date = today
            name1 = Matters[Matters.rfind('被繼承人'):Matters.rfind('被繼承人')+10]
            name2 = Matters[Matters.find('被繼承人'):Matters.find('被繼承人')+10]
            #print(name,name1,name2,Matters)

            if  name in name1 : #'拋棄繼承' in item and

                if '死亡' in Matters.replace(' ','').replace('\n','').replace(' ','') :
                    Matters = Matters.replace(' ','').replace('\n','').replace(' ','')
                    data_1 = Matters[Matters.rfind('死亡')-10:Matters.rfind('死亡')].replace('，','').replace('國','').replace('民','').replace('號','').replace('於','').replace(')','').replace('(','').replace('）','').replace('（','').replace('宣','').replace('告','').replace('、','').replace('生','').replace('死','').replace('亡','')
                    print(Matters)
                    if '年' in data_1 and '月' in data_1 :
                        note2 = {"status":"死亡","Death_date":data_1}
                        note2 = json.dumps(note2,ensure_ascii=False)
                    else:
                        note2 = {"status":"死亡","Death_date":""}
                        note2 = json.dumps(note2,ensure_ascii=False)
                else:
                    note2 = {"status":"死亡","Death_date":""}
                    note2 = json.dumps(note2,ensure_ascii=False)
            elif name in name2 : #'拋棄繼承' in item and

                if '死亡' in Matters.replace(' ','').replace('\n','').replace(' ','') :
                    Matters = Matters.replace(' ','').replace('\n','').replace(' ','')
                    data_1 = Matters[Matters.rfind('死亡')-10:Matters.rfind('死亡')].replace('，','').replace('國','').replace('民','').replace('號','').replace('於','').replace(')','').replace('(','').replace('）','').replace('（','').replace('宣','').replace('告','').replace('、','').replace('生','').replace('死','').replace('亡','')
                    print(Matters)
                    if '年' in data_1 and '月' in data_1 :
                        note2 = {"status":"死亡","Death_date":data_1}
                        note2 = json.dumps(note2,ensure_ascii=False)
                    else:
                        note2 = {"status":"死亡","Death_date":""}
                        note2 = json.dumps(note2,ensure_ascii=False)
                else:
                    note2 = {"status":"死亡","Death_date":""}
                    note2 = json.dumps(note2,ensure_ascii=False)
            else:
                note2 = {"status":"","Death_date":""}
                note2 = json.dumps(note2,ensure_ascii=False)
            #print(Matters[Matters.rfind('死亡')-10:Matters.rfind('死亡')])
            #note = 'Y'
            #flag = '查無資料'
            Matters = Matters[:4000]
            docs = (name,ID,announcement,post_date,register_no,keynote,Basis,Matters,update_date,note,item,note2)
            judicial_result = judicial(docs)
            print(judicial_result)
            if len(rowid) == 0 :
            
                toSQL(judicial_result, totb, server, database, username, password)
                exit_o = exit_obs(server,username,password,database,totb)
                if exit_o >= 100000:
                    sys.exit()
                else:
                    pass
            else:
                delete(server,username,password,database,totb,note,ID,rowid,register_no,item)
                toSQL(judicial_result, totb, server, database, username, password)
                exit_o = exit_obs(server,username,password,database,totb)
                if exit_o >= 100000:
                    sys.exit()
                else:
                    pass
                
            
            
            continue
   
if obs > 0 :
    errormail(obs)

elif obs ==0 :
    errormail(obs)
