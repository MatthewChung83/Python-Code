from bs4 import BeautifulSoup
import requests
import time
import sys
from etl_func import *
from config import *
from dict import *

# 參數設定-資料庫/名單/網站
server,database,username,password,totb1 = db['server'],db['database'],db['username'],db['password'],db['totb1']
entitytype = db['entitytype']
url = wbinfo['url']

obs = src_obs(server,username,password,database,totb1,entitytype)

for i in foo(-1,obs-1):
    
    src = dbfrom(server,username,password,database,totb1,entitytype)
    
    name = src[0][1]
    ID = src[0][2]
    year = src[0][5].replace('年','')
    official_character = src[0][6].replace('字','')
    official_number = src[0][7].replace('號','')
    
    # 逐筆債務人資訊帶入 request post form data
    data = {
        'mall_county_cd':'',
        'mall_name': f'{name}',
        'mall_exam_cer_wn_year': f'{year}',
        'mall_exam_cer_wn_word': f'{official_character}',
        'mall_exam_cer_wn_no': f'{official_number}',
        'mall_idn': f'{ID}',
        'cer_type': '0',
        'jstat': '1',
        }
    
    resp = requests.post(url,data=data)
    soup = BeautifulSoup(resp.text,"lxml")
    #print(soup)
    # 設定回傳預設值-預設查不到相關結果
    status = 'N'
    EAvaliddate = ''
    estateAgent_inc = ''
    updatetime = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))  
    
    if soup.find(action = '/RentQuery/pgp_1_1_1l'):
        tdbody = soup.find(class_ = 'table').find_all('tr')[1].find_all('td')
        EAvaliddate = tdbody[4].text
        estateAgent_inc = tdbody[5].text
        status = 'Y'
        updateSQL(server,username,password,database,totb1,entitytype,estateAgent_inc,EAvaliddate,status,updatetime,ID,name)
        #計算查找筆數，今日調閱大於五千筆則停止     
        exit_o = exit_obs(server,username,password,database,totb1)
        if exit_o >= 5000:
            try:
                driver.close()
            except:
                pass
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
            try:
                driver.close()
            except:
                pass
            sys.exit()
    
    else:
        
        print(entitytype,status,updatetime,ID,name)
        print(estateAgent_inc)
        print(EAvaliddate)
        #print(estateAgent_inc)
        updateSQL(server,username,password,database,totb1,entitytype,estateAgent_inc,EAvaliddate,status,updatetime,ID,name)
        #計算查找筆數，今日調閱大於五千筆則停止     
        exit_o = exit_obs(server,username,password,database,totb1)
        if exit_o >= 5000:
            try:
                driver.close()
            except:
                pass
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
            try:
                driver.close()
            except:
                pass
            sys.exit()
        
    print(ID,status,EAvaliddate,estateAgent_inc)

