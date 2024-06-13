# Browser Driver
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

# Acquire content from html or xml structure 
from urllib.request import urlopen 
from bs4 import BeautifulSoup
import requests
import time
import csv

def w_alert():
    try:
        WebDriverWait(driver, 3).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
    except TimeoutException:
        pass
    
def LoopGetList(list_str,list_id):
    global list_check
    list_check = ''
    while list_check != list_str:
        list_html = BeautifulSoup(driver.page_source, 'html.parser')
        list_check = list_html.find(id=list_id).find_all('option')[0].text
        list_content = list_html.find(id=list_id).find_all('option')
    return list_html

# Pre-Parameter
DriverPath = 'C:\\python_web_driver\\chrome\\chromedriver_win32\\chromedriver'
Url = 'http://easymap.land.moi.gov.tw/R02/Index'

# Call Web Browser and visit specificed url
chromedriver = DriverPath
driver = webdriver.Chrome(chromedriver)
driver.get(Url)
driver.implicitly_wait(30)

# 指向門牌
s_object_01=driver.find_element_by_id('button_addr')
a_object_01=s_object_01.click()
driver.implicitly_wait(30)

# 指向戶政門牌
select = Select(driver.find_element_by_id('doorPlateTypeId'))
select.select_by_index(1)
driver.implicitly_wait(30)

city_list = LoopGetList('縣市','select_city_id1').find(id='select_city_id1').find_all('option')

#for c in range(len(city_list)):
for c in range(1,len(city_list)):
    Select(driver.find_element_by_id('select_city_id1')).select_by_visible_text(city_list[c].text)
    w_alert()
    town_list = LoopGetList('鄉鎮市區','select_town_id1').find(id='select_town_id1').find_all('option')
    
    for t in range(1,len(town_list)):
        Select(driver.find_element_by_id('select_town_id1')).select_by_visible_text(town_list[t].text)
        w_alert()
        road_list = LoopGetList('道路','select_road_id').find(id='select_road_id').find_all('option')
        
        for r in range(1,len(road_list)):
            dataList = [city_list[c].text,town_list[t].text,road_list[r].text,c,t,r]
            csvFileToWrite = open(r'.\address_index.csv', 'a', encoding='utf-8-sig', newline='')
            csvDataToWrite = csv.writer(csvFileToWrite)
            csvDataToWrite.writerow(dataList)
            csvFileToWrite.close()
            print(dataList)