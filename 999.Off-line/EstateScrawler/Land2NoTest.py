import os
import csv
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup

add_src = pd.read_csv(r'C:\Py_Project\project\LandMoi\testing_ds_output.csv').drop_duplicates()

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

    from bs4 import BeautifulSoup
    from requests import Session
    import json
    from fake_useragent import UserAgent

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

import csv
import json
import pandas as pd
import time
import random

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup

#CreateCityIndex()

add_src = pd.read_csv(r'C:\Py_Project\project\LandMoi\testing_ds_output.csv').drop_duplicates()
add_src = add_src.loc[add_src["sync_flag"] != 0].reset_index()

ct_src = pd.read_csv(r'C:\Py_Project\project\LandMoi\city_index.csv').drop_duplicates()
CsvToPath = r'C:\Py_Project\output\LandNo.csv'

LandDataTitle = ('i','csaei','ID','LandNo_f','LandNo_e','District','GeoOffice','Lot','LandNo','AreaSq','LdValue','LdPrice','city','town','road','lane','alley','no')

ReadCsv(CsvToPath,LandDataTitle)

target_file = r'C:\Py_Project\output\LandNo.csv'

if os.path.isfile(target_file):    
    add_out = pd.read_csv(r'C:\Py_Project\output\LandNo.csv').query("i != 'i'").drop_duplicates()
    start = len(add_out)
else:
    start = 0

for i in range(start,len(add_src)):
#for i in range(151,152):
    
    casei = DeColumn(i,'casei',add_src)
    ID = DeColumn(i,'ID',add_src)
    city_in = DeColumn(i,'city',add_src)
    town_in = DeColumn(i,'town',add_src)
    road_in = DeColumn(i,'road',add_src)
    lane_in = DeColumn(i,'lane',add_src)
    alley_in = DeColumn(i,'alley',add_src)
    no_in = DeColumn(i,'no',add_src)
    city_code = ct_src.loc[ct_src["CityName"] == city_in]['CityCode']

    CoordXYUrl = 'http://easymap.land.moi.gov.tw/R02/Door_json_getCoordXY'
    CoordXY = GetCoordXY(CoordXYUrl,city_in,road_in,lane_in,alley_in,no_in,town_in)

    coordX = json.loads(CoordXY.text)['X']
    coordY = json.loads(CoordXY.text)['Y']

    #coordX = 120.362022 #測試用
    #coordY = 22.628284  #測試用

    DoorInfoUrl = 'http://easymap.land.moi.gov.tw/R02/Door_json_getDoorInfoByXY'
    DoorInfoByXY = GetDoorInfoByXY(DoorInfoUrl,city_code,coordX,coordY)

    if '錯誤' in DoorInfoByXY.text:
        LandData = (i,casei,ID,'系統已達上限','','-','-','-','','-','-','-',city_in,town_in,road_in,lane_in,alley_in,no_in)
        ReadCsv(CsvToPath,LandData)
        print(LandData)
    
    elif DoorInfoByXY.text == 'null':
        LandData = (i,casei,ID,'查無地號','查無地號','-','-','-','查無地號','-','-','-',city_in,town_in,road_in,lane_in,alley_in,no_in)
        ReadCsv(CsvToPath,LandData)
        print(LandData)

    else:
        Area = json.loads(DoorInfoByXY.text)['Area']
        towncode = json.loads(DoorInfoByXY.text)['towncode']
        sectno = json.loads(DoorInfoByXY.text)['sectno']
        sectName = json.loads(DoorInfoByXY.text)['sectName']
        landno = json.loads(DoorInfoByXY.text)['landno']
        office = json.loads(DoorInfoByXY.text)['office']
        
        LandDescUrl = 'http://easymap.land.moi.gov.tw/R02/LandDesc_ajax_detail'
        LandDesc = GetLandDesc(LandDescUrl,city_code,towncode,office,sectno,landno)

        if '查無此宗地屬性' in LandDesc.text:
        
            LandData = (i,casei,ID,'查無地號','查無地號','-','-','-','查無地號','-','-','-',city_in,town_in,road_in,lane_in,alley_in,no_in)
            ReadCsv(CsvToPath,LandData)
            print(LandData)

        elif '錯誤' in LandDesc.text:
            LandData = (i,casei,ID,'系統已達上限','','-','-','-','','-','-','-',city_in,town_in,road_in,lane_in,alley_in,no_in)
            #ReadCsv(CsvToPath,LandData)
            print(LandData)

        else:
            District = LandDesc.select('td')[0].text
            GeoOffice = LandDesc.select('td')[1].text
            Lot = LandDesc.select('td')[2].text
            LandNo = LandDesc.select('td')[3].text
            LandNo_f = LandDesc.select('td')[3].text[:4]
            LandNo_e = LandDesc.select('td')[3].text[4:]
            AreaSq = LandDesc.select('td')[4].text
            LdValue = LandDesc.select('td')[5].text
            LdPrice = LandDesc.select('td')[6].text

            LandData = (i,casei,ID,LandNo_f,LandNo_e,District,GeoOffice,Lot,LandNo,AreaSq,LdValue,LdPrice,city_in,town_in,road_in,lane_in,alley_in,no_in)
            ReadCsv(CsvToPath,LandData)
            print(LandData)
    time.sleep(0.5)