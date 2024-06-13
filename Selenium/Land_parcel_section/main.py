from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from requests.exceptions import Timeout
from requests import exceptions
from selenium import webdriver
from bs4 import BeautifulSoup

from PIL import Image
import urllib
import requests
import time
import re
import io

# import import_ipynb
from etl_func import *
from config import *

server,database,username,password,totb = db['server'],db['database'],db['username'],db['password'],db['totb']
Drvfile,url,acc_name,acc_pwd,respimg = wbinfo['Drvfile'],wbinfo['url'],wbinfo['acc_name'],wbinfo['acc_pwd'],wbinfo['respimg']
q = 0
truncate(server,username,password,database)

chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])
driver = webdriver.Chrome(Drvfile,options=chrome_options)
driver.get(url)

driver.find_element(By.ID,'ok').click()
driver.find_element(By.ID,'yes').click()
driver.find_element(By.ID,'Img2').click()

capcha_resp(driver,respimg,q,acc_name,acc_pwd)
driver.switch_to.default_content()
driver.switch_to.frame('untie')
driver.find_element(By.NAME,'Image14').click()

driver.switch_to.default_content()
driver.switch_to.frame('untie')
soup_c = BeautifulSoup(driver.page_source,'html.parser')

city_areas = []

for c in soup_c.find(id = 'City_ID').find_all('option'):
    if c['value'] != '':
        driver.switch_to.default_content()
        driver.switch_to.frame('untie')
        select = Select(driver.find_element(By.ID,'City_ID'))
        select.select_by_value(c['value'])

        soup_a = BeautifulSoup(driver.page_source,'html.parser')
        for a in soup_a.find(id = 'area_id').find_all('option'):
            if a['value'] != '':
                city_area = [c['value'],c.text,a['value'],a.text]
                if city_area not in city_areas:
                    city_areas.append(city_area)

        driver.back()
        time.sleep(1)
driver.close()

for ca in range(len(city_areas)):

    driver = webdriver.Chrome(Drvfile,options=chrome_options)
    driver.get(url)
    driver.find_element(By.ID,'ok').click()
    driver.find_element(By.ID,'yes').click()
    driver.find_element(By.ID,'Img2').click()

    capcha_resp(driver,respimg,q,acc_name,acc_pwd)
    driver.switch_to.default_content()
    driver.switch_to.frame('untie')
    driver.find_element(By.NAME,'Image14').click()

    driver.switch_to.default_content()
    driver.switch_to.frame('untie')
    soup_c = BeautifulSoup(driver.page_source,'html.parser')

    select = Select(driver.find_element(By.ID,'City_ID'))
    select.select_by_value(city_areas[ca][0])
    soup_a = BeautifulSoup(driver.page_source,'html.parser')

    select = Select(driver.find_element(By.ID,'area_id'))
    select.select_by_value(city_areas[ca][2])
    time.sleep(3)
    
    soup_s = BeautifulSoup(driver.page_source,'html.parser')
        
    for s in soup_s.find_all(class_ = 'session_name'):
        if s['id'] != '':
            updatetime = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            doc = (city_areas[ca][0],city_areas[ca][1],city_areas[ca][2],city_areas[ca][3],s['id'],s['name'],is_all_chinese(s['name']),updatetime)
            docs = indata(doc)
            toSQL(docs, totb, server, database, username, password)
            print(doc)
    driver.close()
overwrite(server,username,password,database)
print('更新完成')