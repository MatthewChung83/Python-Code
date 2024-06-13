from bs4 import BeautifulSoup
import requests
import datetime
import time
import re
import sys


from etl_func import *
from config import *
from dict import *

# 宣告資料庫、網站、迭代...等參數
server,database,username,password,fromtb,totb = db['server'],db['database'],db['username'],db['password'],db['fromtb'],db['totb']
url,captchaImg = wbinfo['url'],wbinfo['captchaImg']
imgp = pics['imgp']
q = 0

# 取名單數量
obs = src_obs(server,username,password,database,fromtb,totb)

for i in foo(-1,obs-1):
    # for i in range(0,1):
    src = dbfrom(server,username,password,database,fromtb,totb)
    personi = src[0][0]
    id = src[0][1]
    print(id)
    name = src[0][2]#.encode('latin1').decode('gbk',errors = 'ignore')
    birdte_ad = str(src[0][3]).replace('-','/')
    birdte_ad_tw = '0'+src[0][4].replace('-','/')
    resp = capcha_resp(url,captchaImg,imgp,q,name,id,birdte_ad,birdte_ad_tw)
    #print(resp)
    
    updatetime = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    status,item,occupation,level,license_no,effective_date,remark,M = '','','','','','','',birdte_ad_tw
    
    if '對不起，您目前無技術士證資料' in str(resp.find('script', type='text/javascript')):
        status = 'N'
        delete(server,username,password,database,totb,id,personi)
        doc = (personi,id,name,M,item,occupation,level,license_no,effective_date,remark,status,updatetime)
        docs = indata(doc)
        toSQL(docs, totb, server, database, username, password)
        
    else:
        status = 'Y'
        
        for j in range(1,len(resp.find_all('tr'))):
            row,remark = [],[]
            
            for k in range(len(resp.find_all('tr')[j].find_all('td'))):
                text = resp.find_all('tr')[j].find_all('td')[k].text
                row.append(text)
                print(row)
                
            remark = '職稱:' + row[1] + '_級別:' + row[2] +'_證照編號:'+str(row[3])+'_生效日期:' + row[4] 
            row.append(remark)
            delete(server,username,password,database,totb,id,personi)
            doc = (personi,id,name,M,row[0],row[1],row[2],row[3],row[4],row[5],status,updatetime)
            docs = indata(doc)
            toSQL(docs, totb, server, database, username, password)
    
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