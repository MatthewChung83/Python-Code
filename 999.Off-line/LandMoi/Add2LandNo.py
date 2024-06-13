import csv
import json
import pandas as pd
import time
import random

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup

import PreAct

#PreAct.CreateCityIndex()

add_src = pd.read_csv(r'C:\Py_Project\output\address_sync_output.csv').drop_duplicates()
add_src = add_src.loc[add_src["sync_flag"] != 0].reset_index()

ct_src = pd.read_csv(r'.\city_index.csv').drop_duplicates()
CsvToPath = r'C:\Users\cyche\Documents\Python Scripts\land_moi\LandNo.csv'

LandDataTitle = ('i','csaei','ID','LandNo_f','LandNo_e','District','GeoOffice','Lot','LandNo','AreaSq','LdValue','LdPrice','city','town','road','lane','alley','no')

PreAct.ReadCsv(CsvToPath,LandDataTitle)

#for i in range(len(add_src)):
for i in range(151,152):
    
    casei = PreAct.DeColumn(i,'casei',add_src)
    ID = PreAct.DeColumn(i,'ID',add_src)
    city_in = PreAct.DeColumn(i,'city',add_src)
    town_in = PreAct.DeColumn(i,'town',add_src)
    road_in = PreAct.DeColumn(i,'road',add_src)
    lane_in = PreAct.DeColumn(i,'lane',add_src)
    alley_in = PreAct.DeColumn(i,'alley',add_src)
    no_in = PreAct.DeColumn(i,'no',add_src)
    city_code = ct_src.loc[ct_src["CityName"] == city_in]['CityCode']

    CoordXYUrl = 'http://easymap.land.moi.gov.tw/R02/Door_json_getCoordXY'
    CoordXY = PreAct.GetCoordXY(CoordXYUrl,city_in,road_in,lane_in,alley_in,no_in,town_in)

    coordX = json.loads(CoordXY.text)['X']
    coordY = json.loads(CoordXY.text)['Y']

    #coordX = 120.362022 #測試用
    #coordY = 22.628284  #測試用

    DoorInfoUrl = 'http://easymap.land.moi.gov.tw/R02/Door_json_getDoorInfoByXY'
    DoorInfoByXY = PreAct.GetDoorInfoByXY(DoorInfoUrl,city_code,coordX,coordY)

    if '錯誤' in DoorInfoByXY.text:
        LandData = (i,casei,ID,'系統已達上限','','-','-','-','','-','-','-',city_in,town_in,road_in,lane_in,alley_in,no_in)
        PreAct.ReadCsv(CsvToPath,LandData)
        print(LandData)
    
    elif DoorInfoByXY.text == 'null':
        LandData = (i,casei,ID,'查無地號','查無地號','-','-','-','查無地號','-','-','-',city_in,town_in,road_in,lane_in,alley_in,no_in)
        PreAct.ReadCsv(CsvToPath,LandData)
        print(LandData)

    else:
        Area = json.loads(DoorInfoByXY.text)['Area']
        towncode = json.loads(DoorInfoByXY.text)['towncode']
        sectno = json.loads(DoorInfoByXY.text)['sectno']
        sectName = json.loads(DoorInfoByXY.text)['sectName']
        landno = json.loads(DoorInfoByXY.text)['landno']
        office = json.loads(DoorInfoByXY.text)['office']
        
        LandDescUrl = 'http://easymap.land.moi.gov.tw/R02/LandDesc_ajax_detail'
        LandDesc = PreAct.GetLandDesc(LandDescUrl,city_code,towncode,office,sectno,landno)

        if '查無此宗地屬性' in LandDesc.text:
        
            LandData = (i,casei,ID,'查無地號','查無地號','-','-','-','查無地號','-','-','-',city_in,town_in,road_in,lane_in,alley_in,no_in)
            PreAct.ReadCsv(CsvToPath,LandData)
            print(LandData)

        elif '錯誤' in LandDesc.text:
            LandData = (i,casei,ID,'系統已達上限','','-','-','-','','-','-','-',city_in,town_in,road_in,lane_in,alley_in,no_in)
            PreAct.ReadCsv(CsvToPath,LandData)
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
            PreAct.ReadCsv(CsvToPath,LandData)
            print(LandData)
    time.sleep(5)