import os
import csv
import json
import time
import pymssql
import pandas as pd
from requests import Session
from fake_useragent import UserAgent
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup

db_server = 'dnt79'
db_acct = 'pyuser'
db_pwd = 'Ucredit7607'
wk_tb = 'ElAddrTblandNo79'
df_tb = 'ElAddrTblandNo'
id_start = 0
id_end = 400000

# 在post前清洗 city,town,road,lane,alley,no 的資料格式
def DeColumn(i,ColName,src):
    if src[ColName][i] == src[ColName][i]:
        text = src[ColName][i]
    else:
        text = ''
    return text

# 取得地址的經緯度座標
def GetCoordXY(url,cityName,road,lane,alley,no,area):

    from bs4 import BeautifulSoup
    from requests import Session
    import json
    from fake_useragent import UserAgent
    
    rlan = road+lane+alley+no
    # Call CoordXY
    
    # tor
    proxies = {'http': 'socks5://127.0.0.1:9150','https': 'socks5://127.0.0.1:9150'}
    # tor

    req_session = Session()
    CoordXY_url = url
    resp = req_session.post(CoordXY_url)
    # tor
    # resp = req_session.post(CoordXY_url,proxies=proxies)
    # tor
    cookie = '; '.join( [ '='.join(i) for i in resp.cookies.items()] )
    uar = UserAgent().random

    headers={
        'Cookie': cookie,
        'Host': 'easymap.land.moi.gov.tw',
        'Origin': 'http://easymap.land.moi.gov.tw',
        'Referer': 'http://easymap.land.moi.gov.tw/R02/Index',
        'User-Agent': uar,
        #'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36',
    }
    
    # get token
    set_token_url = 'http://easymap.land.moi.gov.tw/R02/pages/setToken.jsp'
    token_resp = req_session.post(set_token_url, headers=headers)
    # tor
    # token_resp = req_session.post(set_token_url, headers=headers,proxies = proxies)
    # tor
    token = BeautifulSoup(token_resp.text, 'lxml').select('input[name=token]')[0]['value']
    
    data = {
        'cityName': cityName,
        'doorPlate': rlan,
        'townName': area,
        'struts.token.name': 'token',
        'token': token,
    }
    resp=req_session.post(CoordXY_url, headers=headers, data=data,)
    # tor    
    # resp=req_session.post(CoordXY_url, headers=headers,proxies = proxies, data=data)
    # tor
    CoordXY = BeautifulSoup(resp.text,'html.parser')
    return CoordXY

    # 取得地址的地政資訊
def GetDoorInfoByXY(url,City_Code,coordX,coordY):

    # tor
    proxies = {'http': 'socks5://127.0.0.1:9150','https': 'socks5://127.0.0.1:9150'}
    # tor
        
    # Call DoorInfoByXY
    req_session = Session()
    DoorInfoByXY_url = url
    resp = req_session.post(DoorInfoByXY_url)
    # tor
    # resp = req_session.post(DoorInfoByXY_url,proxies=proxies)
    # tor
    cookie = '; '.join( [ '='.join(i) for i in resp.cookies.items()] )

    uar = UserAgent().random

    headers={
        'Cookie': cookie,
        'Host': 'easymap.land.moi.gov.tw',
        'Origin': 'http://easymap.land.moi.gov.tw',
        'Referer': 'http://easymap.land.moi.gov.tw/R02/Index',
        'User-Agent': uar,
    }
    
    # get token
    set_token_url = 'http://easymap.land.moi.gov.tw/R02/pages/setToken.jsp'
    token_resp = req_session.post(set_token_url, headers=headers)
    # tor
    # token_resp = req_session.post(set_token_url, headers=headers,proxies = proxies)
    # tor
    token = BeautifulSoup( token_resp.text, 'lxml').select('input[name=token]')[0]['value']
    
    data = {
        'city': City_Code,
        'coordX': coordX,
        'coordY': coordY,
        'struts.token.name': 'token',
        'token': token,
    }
    resp=req_session.post(DoorInfoByXY_url, headers=headers, data=data)
    # tor    
    # resp=req_session.post(DoorInfoByXY_url, headers=headers,proxies = proxies, data=data)
    # tor
    DoorInfoByXY = BeautifulSoup(resp.text,'html.parser')
    return DoorInfoByXY

# 取得地址的地號
def GetLandDesc(url,City_Code,towncode,office,sectno,landno):

    from bs4 import BeautifulSoup

    from requests import Session
    import json
    from fake_useragent import UserAgent
    
    # Call LandDesc
    # tor
    proxies = {'http': 'socks5://127.0.0.1:9150','https': 'socks5://127.0.0.1:9150'}
    # tor
    req_session = Session()
    LandDesc_url = url
    resp = req_session.post(LandDesc_url)
    # resp = req_session.post(LandDesc_url,proxies=proxies)
    cookie = '; '.join( [ '='.join(i) for i in resp.cookies.items()] )

    uar = UserAgent().random

    headers={
        'Cookie': cookie,
        'Host': 'easymap.land.moi.gov.tw',
        'Origin': 'http://easymap.land.moi.gov.tw',
        'Referer': 'http://easymap.land.moi.gov.tw/R02/Index',
        'User-Agent': uar,
        #'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.106 Safari/537.36',
    }
    
    # get token
    set_token_url = 'http://easymap.land.moi.gov.tw/R02/pages/setToken.jsp'
    token_resp = req_session.post(set_token_url, headers=headers)
    # token_resp = req_session.post(set_token_url, headers=headers,proxies = proxies)
    token = BeautifulSoup( token_resp.text, 'lxml').select('input[name=token]')[0]['value']
    
    data = {
        'cityCode': City_Code,
        'townCode': towncode,
        'office': office,
        'sectNo': sectno,
        'landNo': landno,
        'struts.token.name': 'token',
        'token': token,
    }

    # LandDesc result
    resp=req_session.post(LandDesc_url, headers=headers, data=data)
    # resp=req_session.post(LandDesc_url, headers=headers,proxies = proxies, data=data)
    LandDesc = BeautifulSoup(resp.text,'html.parser')
    return LandDesc

# 寫入的csv檔案
def ReadCsv(file_path,datalist):
    csvFileToWrite = open(file_path, 'a', encoding='utf-8-sig', newline='')
    csvDataToWrite = csv.writer(csvFileToWrite)
    csvDataToWrite.writerow(datalist)
    csvFileToWrite.close()
    
def isrecexist(table,pname,pvalue):
    conn = pymssql.connect(server=db_server, user=db_acct, password=db_pwd)
    cursor = conn.cursor()
    script = f"SELECT * FROM [UIS].[dbo].[{table}] WHERE [{pname}] = {pvalue}"
    cursor.execute(script)
    query = cursor.fetchall()
    cursor.close()
    conn.close()
    return len(query)

def inserttodb(table,eaid,pid,addressi,casei,LandNo_f,LandNo_b,District,GeoOffice,Lot,LandNo,AreaSq,LdValue,LdPrice,datatime,result):
    conn = pymssql.connect(server=db_server, user=db_acct, password=db_pwd)
    cursor2 = conn.cursor()
    script = f"""
    INSERT INTO [UIS].[dbo].[{table}] ([eaid],[pid],[addressi],[casei],[LandNo_f],[LandNo_b],[District],[GeoOffice],[Lot],[LandNo],[AreaSq],[LdValue],[LdPrice],[datatime],[result]) 
    VALUES ({eaid},'{pid}',{addressi},'{casei}','{LandNo_f}','{LandNo_b}','{District}','{GeoOffice}','{Lot}','{LandNo}','{AreaSq}','{LdValue}','{LdPrice}','{datatime}','{result}')
    """
    cursor2.execute(script)
    conn.commit()
    cursor2.close()
    conn.close()
    
def restart():
    os.chdir("C:/Py_Project/script/")
    os.startfile("win-redial.bat")

    # 資料來源導入
ct_src = pd.read_csv(r'C:\Py_Project\project\LandMoi\city_index.csv').drop_duplicates()

def src_in(table,pname,offobs,nextobs):
    conn = pymssql.connect(server=db_server, user=db_acct, password=db_pwd)
    cursor = conn.cursor()
    script = f"""
    SELECT * FROM [UIS].[dbo].[{table}] 
    WHERE result in ('Y') and [{pname}] between {id_start} and {id_end}
    and [{pname}] > (select top 1 [{pname}]  from [UIS].[dbo].[{wk_tb}] order by [{pname}]  desc)
	and [{pname}] not in (select [{pname}] from [UIS].[dbo].[{df_tb}] union select [{pname}] from [UIS].[dbo].[{wk_tb}])
    ORDER BY [{pname}]
    offset {offobs} row fetch
    next {nextobs} rows only
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src
    
c_src = src_in('ElAddrTbCheck','eaid',0,200000)

for i in range(len(c_src)):

    eaid = c_src[i][0]
    ID = c_src[i][1]
    addressi = c_src[i][2]
    casei = c_src[i][3]
    city_in = c_src[i][4]
    town_in = c_src[i][5]
    road_in = c_src[i][6]
    lane_in = c_src[i][7]
    alley_in = c_src[i][8]
    no_in = c_src[i][9]+c_src[i][10]
    datatime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    city_code = ct_src.loc[ct_src["CityName"] == city_in]['CityCode']
    
    if isrecexist(wk_tb,'eaid',eaid) > 0 and i != len(c_src) - 1:
        continue
    elif isrecexist(wk_tb,'eaid',eaid) > 0 and i == len(c_src) - 1:
        break
    else:
        pass
    
    CoordXYUrl = 'http://easymap.land.moi.gov.tw/R02/Door_json_getCoordXY'
    CoordXY = GetCoordXY(CoordXYUrl,city_in,road_in,lane_in,alley_in,no_in,town_in)

    if '網頁發生錯誤' in CoordXY.text:
        print(i,eaid,'CoordXY',CoordXY.text.strip(),time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        District = ''
        GeoOffice = ''
        Lot = ''
        LandNo = ''
        LandNo_f = ''
        LandNo_b = ''
        AreaSq = ''
        LdValue = ''
        LdPrice = ''
        result = 'N'
        inserttodb(wk_tb,eaid,ID,addressi,casei,LandNo_f,LandNo_b,District,GeoOffice,Lot,LandNo,AreaSq,LdValue,LdPrice,datatime,result)
        continue
        
    elif '系統發生錯誤' in CoordXY.text or 'ACCESS DENY' in CoordXY.text:
        print(i,eaid,'CoordXY',CoordXY.text.strip(),time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        restart()
        time.sleep(20)
    else:
        coordX = json.loads(CoordXY.text)['X']
        coordY = json.loads(CoordXY.text)['Y']
        DoorInfoUrl = 'http://easymap.land.moi.gov.tw/R02/Door_json_getDoorInfoByXY'
        DoorInfoByXY = GetDoorInfoByXY(DoorInfoUrl,city_code,coordX,coordY)
        
        if '網頁發生錯誤' in DoorInfoByXY.text:
            print(i,eaid,'DoorInfoByXY',DoorInfoByXY.text.strip(),time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            District = ''
            GeoOffice = ''
            Lot = ''
            LandNo = ''
            LandNo_f = ''
            LandNo_b = ''
            AreaSq = ''
            LdValue = ''
            LdPrice = ''
            result = 'N'
            inserttodb(wk_tb,eaid,ID,addressi,casei,LandNo_f,LandNo_b,District,GeoOffice,Lot,LandNo,AreaSq,LdValue,LdPrice,datatime,result)
            continue
            
        elif '系統發生錯誤' in DoorInfoByXY.text or 'ACCESS DENY' in DoorInfoByXY.text:
            print(i,eaid,'DoorInfoByXY',DoorInfoByXY.text.strip(),time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            restart()
            time.sleep(20)
            
        elif DoorInfoByXY.text == 'null':
            District = ''            
            GeoOffice = ''
            Lot = ''
            LandNo = ''
            LandNo_f = ''
            LandNo_b = ''
            AreaSq = ''
            LdValue = ''
            LdPrice = ''
            result = 'N'
            inserttodb(wk_tb,eaid,ID,addressi,casei,LandNo_f,LandNo_b,District,GeoOffice,Lot,LandNo,AreaSq,LdValue,LdPrice,datatime,result)
            print(i,eaid,'查無地號',time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            time.sleep(1)
        else:
            Area = json.loads(DoorInfoByXY.text)['Area']
            towncode = json.loads(DoorInfoByXY.text)['towncode']
            sectno = json.loads(DoorInfoByXY.text)['sectno']
            sectName = json.loads(DoorInfoByXY.text)['sectName']
            landno = json.loads(DoorInfoByXY.text)['landno']
            office = json.loads(DoorInfoByXY.text)['office']
            
            LandDescUrl = 'http://easymap.land.moi.gov.tw/R02/LandDesc_ajax_detail'
            LandDesc = GetLandDesc(LandDescUrl,city_code,towncode,office,sectno,landno)
            
            if '網頁發生錯誤' in LandDesc.text:
                print(i,eaid,'LandDesc',LandDesc.text.strip(),time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
                District = ''
                GeoOffice = ''
                Lot = ''
                LandNo = ''
                LandNo_f = ''
                LandNo_b = ''
                AreaSq = ''
                LdValue = ''
                LdPrice = ''
                result = 'N'
                inserttodb(wk_tb,eaid,ID,addressi,casei,LandNo_f,LandNo_b,District,GeoOffice,Lot,LandNo,AreaSq,LdValue,LdPrice,datatime,result)
                continue
                
            elif '系統發生錯誤' in LandDesc.text or 'ACCESS DENY' in LandDesc.text:
                print(i,eaid,'LandDesc',LandDesc.text.strip(),time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
                restart()
                time.sleep(20)

            elif '查無此宗地屬性' in LandDesc.text:
                District = ''
                GeoOffice = ''
                Lot = ''
                LandNo = ''
                LandNo_f = ''
                LandNo_b = ''
                AreaSq = ''
                LdValue = ''
                LdPrice = ''
                result = 'N'
                inserttodb(wk_tb,eaid,ID,addressi,casei,LandNo_f,LandNo_b,District,GeoOffice,Lot,LandNo,AreaSq,LdValue,LdPrice,datatime,result)
                print(i,eaid,'查無地號',time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            else:
                District = LandDesc.select('td')[0].text
                GeoOffice = LandDesc.select('td')[1].text
                Lot = LandDesc.select('td')[2].text
                LandNo = LandDesc.select('td')[3].text
                LandNo_f = LandDesc.select('td')[3].text[:4]
                LandNo_b = LandDesc.select('td')[3].text[4:]
                AreaSq = LandDesc.select('td')[4].text
                LdValue = LandDesc.select('td')[5].text
                LdPrice = LandDesc.select('td')[6].text
                result = 'Y'
                inserttodb(wk_tb,eaid,ID,addressi,casei,LandNo_f,LandNo_b,District,GeoOffice,Lot,LandNo,AreaSq,LdValue,LdPrice,datatime,result)
                print(i,eaid,'insert into',time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    time.sleep(0.5)