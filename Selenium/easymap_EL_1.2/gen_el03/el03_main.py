import csv
import sys
import json
import time
import pymssql
import datetime as dt
import argparse
import pandas as pd
import os

from requests import Session
from bs4 import BeautifulSoup
from selenium import webdriver
from fake_useragent import UserAgent
from selenium.webdriver.support.ui import Select


from config import db
from batch_args import *
from sys_env import restart
from etl_func import *
from scrawer import gettowninfo,getdoorinfo,getbuildinfo,GetCoordXY,GetDoorInfoByXY,GetLandDesc

### phase0：driver outer argument ###
server = db['server']
database = db['database']
username = db['username']
password = db['password']
fromtb = db['fromtb']
totb = db['totb']
todtltb = db['todtltb']
entitytype = db['entitytype']
batstoptime = dt.datetime.strptime(db['batstoptime'], '%Y-%m-%d %H:%M:%S.%f')


validtrans = 0
actiondttm = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
Statusdttm = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
Machine = hostname()
EntityPhase = db['EntityPhase']
EntityPhase_next = db['EntityPhase_next']
EntityPath_next = db['EntityPath_next']
todtltb = db['todtltb']
MGtb = db['MGtb']




#分台執行(10.90.0.90)
#conn = pymssql.connect(server=server, user=username, password=password, database = database)
#cursor = conn.cursor()
#script = f"""
#drop Table EL02_02
#SELECT * ,NTILE (4)  OVER (ORDER BY eaid DESC) AS flag
#into EL02_02
#FROM EL02
#where type='{entitytype}'
#"""    
#cursor.execute(script)
#conn.commit()
#cursor.close()
#conn.close()
#

#合併寫入(10.90.0.90)
#conn = pymssql.connect(server=server, user=username, password=password, database = database)
#cursor = conn.cursor()
#script = f"""
#insert into EL03
#select * from EL03_02

#insert into EL03DTL
#select * from EL03DTL_02
#"""    
#cursor.execute(script)
#conn.commit()
#cursor.close()
#conn.close()
#


doorPlateType = 'A'
obs = src_obs(server,username,password,database,fromtb,totb,entitytype)
records = []


for i in foo(-1,obs-1):
    count_res = 1
    while count_res < 5:
        try:
            bd_log,ld_log,result = '','',''
            cityName,area,road,lane,alley,no = '','','','','',''
            src = dbfrom(server,username,password,database,fromtb,totb,entitytype)
            
            eaid,pid,addressi,casei,entitytype = src[0][0],src[0][1],src[0][2],src[0][3],entitytype
            cityName,area,road,lane,alley,no = src[0][4],src[0][5],src[0][6],src[0][7],src[0][8],src[0][9]+src[0][10]
            print(eaid)
            ct_src = pd.read_csv(r'C:\Py_Project\project\LandMoi\city_index.csv').drop_duplicates()
            cityCode = list(ct_src.loc[ct_src["CityName"] == cityName]['CityCode'])[0]
            
            towncode = ''
            doorinfo,buildno,sectno,office = {},'','',''
            bd_district,bd_geooffice,bd_lot,builddno,bd_areasq,bd_floorno,bd_byfloor,bd_completiondate,bd_purpose,bd_landno = '','','','','','','','','',''
            
            CoordX,CoordY = '',''
            towncode,sectno,office,landno='','','',''
            ld_district,ld_geooffice,ld_lot,landno,ld_areasq,ld_value,ld_price,landno_f,landno_b = '','','','','','','','',''
            
        ### phase1：get buildno ###
        
            # 1.1 - get towncode
            getTown_url = 'https://easymap.land.moi.gov.tw/P02/City_json_getTownList'
            try:
                result = gettowninfo(getTown_url,cityCode,cityName,doorPlateType)
            except:
                restart()
                time.sleep(70)
                result = gettowninfo(getTown_url,cityCode,cityName,doorPlateType)
            bd_towninfoLog = result    
        
        
            try:
                towncode = [town for town in json.loads(result) if town['name'] == area][0]['id']
            except:
                msg = '(area 輸入內容不合規 at gettowninfo)'
                towncode = ''
                bd_log = bd_log + msg
            
            # 1.2 - get buildno,sectno,sectName,office,landno    
            if towncode == '':
                pass        
            
            else:
                getDoor_url = 'https://easymap.land.moi.gov.tw/P02/Door_json_getDoorList'
                try:
                    result = getdoorinfo(getDoor_url,cityCode,towncode,road,doorPlateType,lane,alley,area,cityName,no)
                    #print(result)
                except:
                    restart()
                    time.sleep(70)
                    result = getdoorinfo(getDoor_url,cityCode,towncode,road,doorPlateType,lane,alley,area,cityName,no)
                bd_doorinfoLog = result
                try:
                    soup = BeautifulSoup(result, 'html.parser')
                    check = soup.find('title').text
                    print(check)
                    if 'HTTP Status 403 – Forbidden' in check or '網頁發生錯誤' in check :
                        result = 'NN'
                        bd_log = '門牌地址輸入參數錯誤'
                      #print(result)
                        insertdate = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
                        docs = (eaid,pid,addressi,casei,entitytype,
                                bd_district,bd_geooffice,bd_lot,buildno,bd_areasq,bd_floorno,bd_byfloor,bd_completiondate,bd_purpose,
                                ld_district,ld_geooffice,ld_lot,landno,landno_f,landno_b,ld_areasq,ld_value,ld_price,
                                result,insertdate,bd_log,ld_log,bd_landno)
                        print(docs)
                       #print(docs)
                        EL03_result = EL03_etl(docs)
                        toSQL(EL03_result, totb, server, database, username, password)
                        continue
                    else:
                        pass
                except:
                    pass
                try:
                    getTown_results=json.loads(bd_doorinfoLog)['results']
                    #print(getTown_results)
                except:
                    restart()
                    time.sleep(70)
                    pass
                
                
                if '系統發生錯誤' in bd_doorinfoLog or 'ACCESS DENY' in bd_doorinfoLog:
                    msg = '(連線數達到上限 at getdoorinfo)'
                    bd_log = bd_log + msg
            
                elif '發生錯誤' in bd_doorinfoLog:
                    msg = '(網站不明錯誤 at getdoorinfo)'
                    bd_log = bd_log + msg
            
                elif  json.loads(bd_doorinfoLog)['results'] == []:
                    msg = '(本筆資料未辦建物登記 at getdoorinfo)'
                    bd_log = bd_log + msg
            
                else:
                    try:
                        doorinfo = json.loads(result)['results'][0]
                        buildno,sectno,office,bd_landno=doorinfo['buildno'],doorinfo['sectno'],doorinfo['office'],doorinfo['landno']
                    except:
                        restart()
                        time.sleep(70)
                        doorinfo = json.loads(result)['results'][0]
                        buildno,sectno,office,bd_landno=doorinfo['buildno'],doorinfo['sectno'],doorinfo['office'],doorinfo['landno']
                        pass
            # 1.3 - buildinfo
                if '(本筆資料未辦建物登記 at getdoorinfo)' not in bd_log:
                    build_url = 'https://easymap.land.moi.gov.tw/P02/BuildingDesc_ajax_detail'
                    try:
                        result = getbuildinfo(build_url,office,sectno,buildno)
                    except:
                        restart()
                        time.sleep(70)
                        result = getbuildinfo(build_url,office,sectno,buildno)
                    buildinfo = {}
                    doc=BeautifulSoup(result,'lxml')
                    bd_buildinfoLog = doc.text
                    name_list = doc.select('p>small')
                    a=0
                    for name in name_list:
                        a+=1
                        if a==1:
                            buildinfo['行政區']=name.get_text().strip()
                        if a==2:
                            buildinfo['地政事務所']=name.get_text().strip()
                        if a==3:
                            buildinfo['地段']=name.get_text().strip()
                        if a==4:
                            buildinfo['建號']=name.get_text().strip()        
                        if a==5:
                            buildinfo['建物面積']=name.get_text().strip() 
                        if a==6:
                            buildinfo['樓層數']=name.get_text().strip()
                        if a==7:
                            buildinfo['樓層別']=name.get_text().strip()
                        if a==8:
                            buildinfo['建物完成日期']=name.get_text().strip()
                        if a==9:
                            buildinfo['主要用途']=name.get_text().strip()
                    #for row in doc.select('tr'):
                        #row_name = row.th
                        #if row_name:
                            #buildinfo[row_name.text]=row.td.text
        
        
                    if '系統發生錯誤' in doc.text:
                        msg = '(連線數達到上限 at getbuildinfo)'
                        bd_log = bd_log + msg        
            
                    elif '取得建物標示部資料發生錯誤' in doc.text or '發生錯誤' in doc.text:
                        msg = '(取得建物標示部資料發生錯誤 at getbuildinfo)'
                        bd_log = bd_log + msg        
        
                    elif ','.join(buildinfo.values()) == '查無此建物屬性' or json.loads(bd_doorinfoLog)['results'] == []:
                        msg = '(查無此建物屬性 at getbuildinfo)'
                        bd_log = bd_log + msg
                        
                    else:
                        bd_district,bd_geooffice,bd_lot,builddno = buildinfo['行政區'],buildinfo['地政事務所'],buildinfo['地段'],buildinfo['建號']
                        bd_areasq,bd_floorno,bd_byfloor = buildinfo['建物面積'],buildinfo['樓層數'],buildinfo['樓層別']
                        bd_completiondate,bd_purpose = buildinfo['建物完成日期'],buildinfo['主要用途']
        
        ### phase2：get landno ###
        
            # 2.1 - X,Y Coord
            CoordXYUrl = 'https://easymap.land.moi.gov.tw/P02/Door_json_getCoordXY'
            try:
                result = GetCoordXY(CoordXYUrl,cityName,road,lane,alley,no,area)
            except:
                restart()
                time.sleep(70)
                result = GetCoordXY(CoordXYUrl,cityName,road,lane,alley,no,area)
            ld_CoordXYLog = result
        
            try:
                CoordX,CoordY = json.loads(result.text)['X'],json.loads(result.text)['Y']
            except :
                msg = '(missing towncode at phrase 2 step 1 get CoordXY at SYSError)'
                ld_log = ld_log + msg
        
            # 2.2 - get towncode,sectno,office,landno
        
            DoorInfoUrl = 'https://easymap.land.moi.gov.tw/P02/Door_json_getDoorInfoByXY'
            try:
                result = GetDoorInfoByXY(DoorInfoUrl,cityCode,CoordX,CoordY,road,lane,alley,no).text
            except:
                restart()
                time.sleep(70)
                result = GetDoorInfoByXY(DoorInfoUrl,cityCode,CoordX,CoordY,road,lane,alley,no).text
            ld_DoorInfoLog = result
        
            if '系統發生錯誤' in result:
                msg = '(連線數達到上限 at GetDoorInfoByXY)'
                ld_log = ld_log + msg
                
            elif '發生錯誤' in result or ld_DoorInfoLog == 'null':
                msg = '(網站不明錯誤 at GetDoorInfoByXY)'
                ld_log = ld_log + msg
        
            else:
                try:
                    towncode,sectno,office,landno = json.loads(result)['towncode'],json.loads(result)['sectno'],json.loads(result)['office'],json.loads(result)['landno']
                except:
                    restart()
                    time.sleep(70)
                    towncode,sectno,office,landno = json.loads(result)['towncode'],json.loads(result)['sectno'],json.loads(result)['office'],json.loads(result)['landno']
            # 3.3 - get landinfo
            LandDescUrl = 'https://easymap.land.moi.gov.tw/P02/LandDesc_ajax_detail'
            try:
                result = GetLandDesc(LandDescUrl,cityCode,towncode,office,sectno,landno)
            except:
                restart()
                time.sleep(70)
                result = GetLandDesc(LandDescUrl,cityCode,towncode,office,sectno,landno)
            landinfo = {}
            doc=BeautifulSoup(result,'lxml').text
            #ld_landinfoLog = doc.text.replace('使用分區鄉村區\n','').replace('使用地類別乙種建築用地\n','').replace('使用分區特定農業區\n','').replace('使用地類別甲種建築用地\n','').replace('使用地類別農牧用地\n','').replace('使用地類別農牧用地\n','')
            #name_list = ld_landinfoLog.split('\n')
            ld_landinfoLog = doc
            name_list = ld_landinfoLog.split('\n')
            print(eaid)
            
            
            #for row in doc.select('tr'):
                #row_name = row.th
                #if row_name:
                   # landinfo[row_name.text]=row.td.text
        
            if '系統發生錯誤' in doc:
                msg = '(連線數達到上限 at GetLandDesc)'
                ld_log = ld_log + msg
        
            elif '取得土地標示部資料發生錯誤' in doc:
                msg = '(取得土地標示部資料發生錯誤 at GetLandDesc)'
                ld_log = ld_log + msg
                
            elif '查無此宗地屬性' in doc:
                msg = '(查無此宗地屬性 at GetLandDesc)'
                ld_log = ld_log + msg
                
            else:
                try:
                    if '行政區' in name_list[0]:
                        landinfo['行政區']=name_list[0].lstrip('行政區')
                    if '地政事務所' in name_list[1]:
                        landinfo['地政事務所']=name_list[1].lstrip('地政事務所')
                    if '地段' in name_list[2]:
                        landinfo['地段']=name_list[2].lstrip('地段')
                    if '地號' in name_list[3]:
                        landinfo['地號']=name_list[3].lstrip('地號')
                    if '面積' in name_list[4]:
                        landinfo['面積']=name_list[4].lstrip('面積')
                    if '公告土地現值' in name_list[7]:
                        landinfo['公告土地現值']=name_list[7].lstrip('公告土地現值')
                    if '公告土地地價' in name_list[8]:
                        landinfo['公告土地地價']=name_list[8].lstrip('公告土地地價')
                except:
                    if '行政區' in name_list[0]:
                        landinfo['行政區']=name_list[0].lstrip('行政區')
                    if '地政事務所' in name_list[1]:
                        landinfo['地政事務所']=name_list[1].lstrip('地政事務所')
                    if '地段' in name_list[2]:
                        landinfo['地段']=name_list[2].lstrip('地段')
                    if '地號' in name_list[3]:
                        landinfo['地號']=name_list[3].lstrip('地號')
                    if '面積' in name_list[4]:
                        landinfo['面積']=name_list[4].lstrip('面積')
                    if '公告土地現值' in name_list[5]:
                        landinfo['公告土地現值']=name_list[5].lstrip('公告土地現值')
                    if '公告土地地價' in name_list[6]:
                        landinfo['公告土地地價']=name_list[6].lstrip('公告土地地價')
                
                ld_district,ld_geooffice,ld_lot,landno = landinfo['行政區'],landinfo['地政事務所'],landinfo['地段'],landinfo['地號']
                ld_areasq,ld_value,ld_price = landinfo['面積'],landinfo['公告土地現值'],landinfo['公告土地地價']
                landno_f,landno_b = landinfo['地號'][:4],landinfo['地號'][4:]
        
        
        ### phase3：info summarized ###    
            insertdate = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            
            if bd_log == '':
                bd_result = 'Y'
            else:
                bd_result = 'N'
            if ld_log == '':
                ld_result = 'Y'
            else:
                ld_result = 'N'
                
            result = bd_result + ld_result
            if '本筆資料未辦建物登記' in bd_log or '查無此建物屬性' in bd_log:
                docs = (eaid,pid,addressi,casei,entitytype,
                        bd_district,bd_geooffice,bd_lot,buildno,bd_areasq,bd_floorno,bd_byfloor,bd_completiondate,bd_purpose,
                        ld_district,ld_geooffice,ld_lot,landno,landno_f,landno_b,ld_areasq,ld_value,ld_price,
                        result,insertdate,bd_log,ld_log,bd_landno)
            else:
                docs = (eaid,pid,addressi,casei,entitytype,
                        bd_district,bd_geooffice,bd_lot,buildno,bd_areasq,bd_floorno,bd_byfloor,bd_completiondate,bd_purpose,
                        bd_district,bd_geooffice,bd_lot,landno,landno_f,landno_b,ld_areasq,ld_value,ld_price,
                        result,insertdate,bd_log,ld_log,bd_landno)
            
            print(docs)
            #20211119資料回填
            if result == 'YN' :
                conn = pymssql.connect(server=server, user=username, password=password, database = database)
                cursor = conn.cursor()
                script = f"""
                update {totb} 
                set Ld_District=Bd_District , Ld_GeoOffice=Bd_GeoOffice , Ld_Lot=Bd_Lot
                WHERE eaid = eaid and pid = pid ;
                """    
                cursor.execute(script)
                conn.commit()
                cursor.close()
                conn.close()
            
            ### phase4：change IP ###     
            if '連線數達到上限' in bd_log or '連線數達到上限' in ld_log :
                restart()
                time.sleep(70)
            else:
                EL03_result = EL03_etl(docs)
                toSQL(EL03_result, totb, server, database, username, password)
                
                landno_list = landno_split(landno,landno_f,landno_b,bd_landno)
                landno_list = list(set(filter(None,landno_list)))
                
                for n in range(len(landno_list)):
                    text_f = str(int(landno_list[n][0].split('-',-1)[0]) + 10000)[1:]
                    if '-' in landno_list[n][0]:
                        text_b = str(int(landno_list[n][0].split('-',-1)[1]) + 10000)[1:]
                    else:
                        text_b = '0000'
                    flg = landno_list[n][1]
                    if len(ld_district)>0:
                        docs_dtl = (eaid,pid,addressi,entitytype,ld_district,ld_lot,text_f+text_b,flg,insertdate)
                    else :
                        docs_dtl = (eaid,pid,addressi,entitytype,bd_district,bd_lot,text_f+text_b,flg,insertdate)
                    EL03_DTL_result = EL03_DTL_etl(docs_dtl)
                    toSQL(EL03_DTL_result, todtltb, server, database, username, password)
                    
                validtrans = check_obs(server,username,password,database,fromtb,totb,entitytype)
                Statusdttm = insertdate
                end_obs = src_obs(server,username,password,database,fromtb,totb,entitytype)
                
                if i == (obs - 1) and validtrans < obs:
                    status = 'Re-Exec'
                
                elif end_obs == 0:
                    status = 'Done'
                
                else: 
                    status = 'Continue'
                    
                dtlobs = DTLOBS(server,username,password,database,fromtb,totb,entitytype)
                print(dtlobs)
                updateMG_Main(server,username,password,database,validtrans,actiondttm,Statusdttm,status,entitytype,Machine,EntityPhase,MGtb)
                if status == 'Done':
                    status = 'Stand by'
                updateMG_Next(server,username,password,database,actiondttm,Statusdttm,status,entitytype,Machine,EntityPhase_next,EntityPath_next,dtlobs,MGtb,todtltb)
                i+=1
                print(i)
                if i %30 ==0:
                    restart()
                    print('restart')
                    time.sleep(70)
            now=dt.datetime.now()
            
            if now <= batstoptime:
                print('sys exec ; now=',now,' batstoptime=',batstoptime)
                pass
            count_res == 5
        except:
            
            if count_res == 5 :
                result = 'NN'
                bd_log = '門牌地址輸入參數錯誤'
                #print(result)
                insertdate = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
                docs = (eaid,pid,addressi,casei,entitytype,
                        bd_district,bd_geooffice,bd_lot,buildno,bd_areasq,bd_floorno,bd_byfloor,bd_completiondate,bd_purpose,
                        ld_district,ld_geooffice,ld_lot,landno,landno_f,landno_b,ld_areasq,ld_value,ld_price,
                        result,insertdate,bd_log,ld_log,bd_landno)
                print(docs)
                #print(docs)
                EL03_result = EL03_etl(docs)
                toSQL(EL03_result, totb, server, database, username, password)
                obs = src_obs(server,username,password,database,fromtb,totb,entitytype)
                print(obs)
                
                continue
            
            else:
                count_res=count_res+1
                print(count_res)
                quit
        continue
                    #else:
                        #print('sys stop ; now=',now,' batstoptime=',batstoptime)
                        #sys.exit()
                
if obs == 0 :
    SUCC_mail(obs)
    os.system(rf"C:\Py_Project\project\pybatch\02.pybatch_EL_appproc.bat")
    pass
else:
    ERR_mail(obs)
    os.system(rf"C:\Py_Project\project\pybatch\01.pybatch_EL_index.bat")
    pass
