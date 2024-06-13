import csv
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup

add_src = pd.read_csv(r'.\testing_ds_output.csv').drop_duplicates()

# 建立 City and City Code mapping 表
def CreateCityIndex():
    
    DriverPath = 'C:\\python_web_driver\\chrome\\chromedriver_win32\\chromedriver'
    
    Url = 'http://easymap.land.moi.gov.tw/R02/Index'

    file_path = r'C:\Users\cyche\Documents\Python Scripts\land_moi\city_index.csv'
    
    driver = webdriver.Chrome(DriverPath)
    driver.get(Url)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    
    city_index = ('CityCode','CityName')
    csvFileToWrite = open(file_path, 'a', encoding='utf-8-sig', newline='')
    csvDataToWrite = csv.writer(csvFileToWrite)
    csvDataToWrite.writerow(city_index)
    csvFileToWrite.close()

    for tag in driver.find_elements_by_css_selector('select[id=select_city_id] option'):
        value = tag.get_attribute('value')
        text = tag.text
        datalist = [value,text]
        csvFileToWrite = open(file_path, 'a', encoding='utf-8-sig', newline='')
        csvDataToWrite = csv.writer(csvFileToWrite)
        csvDataToWrite.writerow(datalist)
        csvFileToWrite.close()

# 在post前清洗 city,town,road,lane,alley,no 的資料格式
def DeColumn(i,ColName,src):
    if src[ColName][i] == src[ColName][i]:
        text = src[ColName][i]
    else:
        text = ''
    return text

"""
def GetDoorList(url,city,area,road,doorPlate,doorPlateType,lane,alley,townName,cityName,no):
    
    from bs4 import BeautifulSoup
    from requests import Session
    import json
    
    # Call DoorList
    req_session = Session()
    DoorList_url = url    
    resp = req_session.post(DoorList_url)
    cookie = '; '.join( [ '='.join(i) for i in resp.cookies.items()] )

    from fake_useragent import UserAgent
    ua = UserAgent()
    uar = ua.random

    headers={
        'Cookie': cookie,
        'Host': 'easymap.land.moi.gov.tw',
        'Origin': 'http://easymap.land.moi.gov.tw',
        'Referer': 'http://easymap.land.moi.gov.tw/R02/Index',
        'User-Agent': uar,
        #'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.106 Safari/537.36',
    }
    
    set_token_url = 'http://easymap.land.moi.gov.tw/R02/pages/setToken.jsp'
    token_resp = req_session.post(set_token_url, headers=headers)
    token = BeautifulSoup( token_resp.text, 'lxml').select('input[name=token]')[0]['value']
    
    data = {
        'city': city,
        'area': area,
        'road': road,
        'doorPlate': doorPlate,
        'doorPlateType': doorPlateType,
        'lane': lane,
        'alley': alley,
        'townName': townName,
        'cityName': cityName,
        'no': no,
        'struts.token.name': 'token',
        'token': token,
    }
    # DoorList result
    resp=req_session.post(DoorList_url, headers=headers, data=data)
    DoorList=BeautifulSoup(resp.text,'html.parser')
    return DoorList
"""

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
    captchacheck_url = 'http://easymap.land.moi.gov.tw/R02/CaptchaCheck_json_showCheck'
    resp = req_session.post(captchacheck_url)
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

    CoordXY_url = url
    
    # tor    
    resp=req_session.post(CoordXY_url, headers=headers, data=data,)
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
    captchacheck_url = 'http://easymap.land.moi.gov.tw/R02/CaptchaCheck_json_showCheck'
    resp = req_session.post(captchacheck_url)
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
    
    # tor
    token_resp = req_session.post(set_token_url, headers=headers)
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

    DoorInfoByXY_url = url
    
    # tor    
    resp=req_session.post(DoorInfoByXY_url, headers=headers, data=data)
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
    captchacheck_url = 'http://easymap.land.moi.gov.tw/R02/CaptchaCheck_json_showCheck'
    resp = req_session.post(captchacheck_url)
    
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
    LandDesc_url = url
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

# 取Proxy代理
def GetProxyIP(page):
    import requests
    from bs4 import BeautifulSoup
    import json
    from fake_useragent import UserAgent
    import csv
    
    ua = UserAgent()
    uar = ua.random
    url = f'https://www.xicidaili.com/wt/{page}'
    res = requests.get(url, headers={'User-Agent': uar})
    soup = BeautifulSoup(res.text,'html')
    
    #for i in range(1,len(ip_list)):
    for i in range(1,len(ip_list)):
        
        ip_info = ip_list[i].find_all('td')
        
        ip = ip_info[1].text
        port = ip_info[2].text
        country = ip_info[3].text
        level = ip_info[4].text
        segment = ip_info[5].text
        survival = ip_info[8].text
        valid_dt = ip_info[9].text
        
        data = (i,ip,port,country,level,segment,survival,valid_dt)
        print(data)
        
        file_path = r'.\proxy_list.csv'
        csvFileToWrite = open(file_path, 'a', encoding='utf-8-sig', newline='')
        csvDataToWrite = csv.writer(csvFileToWrite)
        csvDataToWrite.writerow(data)
        csvFileToWrite.close()