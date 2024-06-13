# -*- coding: utf-8 -*-
"""
Created on Mon Jan 24 13:11:26 2022

@author: admin
"""


import io
import time
import calendar
import requests
from PIL import Image
#from PIL import ImageGrab
from bs4 import BeautifulSoup
from selenium import webdriver
import selenium.common.exceptions
import pyscreenshot as ImageGrab
import ddddocr

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.expected_conditions import alert_is_present

from config  import *
from dict import *
from etl_func import *





server,database,username,password,totb1,entitytype = db['server'],db['database'],db['username'],db['password'],db['totb1'],db['entitytype']
imgf,imgp = APinfo['imgf'],APinfo['imgp']
url,Drvfile = wbinfo['url'],wbinfo['Drvfile']
driver = webdriver.Chrome(Drvfile)
driver.get(url)
obs = src_obs(server,username,password,database,totb1,entitytype)
for i in foo(-1,obs-1):

    # 逐筆取庫內資料
    src = dbfrom(server,username,password,database,totb1,entitytype)[0]
    name = src[1]
    ID = src[2]
    finall = ''
    driver.find_element(By.ID, 'idnum94').clear()
    driver.find_element(By.ID, 'captchaInput_captcha-refresh').clear()
    time.sleep(1)
    print('A')
        
    driver.find_element(By.ID, 'idnum94').send_keys(ID)
    driver.find_element(By.ID, 'applyTWY').send_keys('94')
    driver.find_element(By.ID, 'applyMM').send_keys('1')
    driver.find_element(By.ID, 'applyDD').send_keys('1')
    driver.find_element(By.ID, 'siteId').send_keys('北縣')
    driver.find_element(By.ID, 'applyReason').send_keys('補發')
    time.sleep(1)
    print('B')
    img_ele = driver.find_element_by_id('captchaImage_captcha-refresh')
    img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
    img.save(imgp)

    ocr = ddddocr.DdddOcr()
    with open(imgp, 'rb') as f:
        img_bytes = f.read()
    res = ocr.classification(img_bytes)
    
    print(res)
    time.sleep(1)
    driver.find_element(By.ID, 'captchaInput_captcha-refresh').send_keys(res)
    time.sleep(1)
    driver.find_element_by_class_name("btn.btn-primary.query").click()
    print('C')
    try:
        finall = driver.find_element(By.ID, 'captchaInput.errors').text
        print('D'+finall)
    except:
        pass
    if '驗證碼' in finall:
        driver.close()
        driver = webdriver.Chrome(Drvfile)
        driver.get(url)
        continue
    try:
        print('E')
        time.sleep(1)
        finall = driver.find_element_by_xpath("//*[@id='resultBlock']/div[2]/div[2]/div[1]/div[2]/div/div/div").text
        #print(finall)
        status = 'Y'
        note = finall
        updatetime = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        updateSQL(server,username,password,database,totb1,entitytype,status,updatetime,ID,name,note)
        print(ID,name,note,status)
        driver.find_element_by_class_name("btn.btn-primary.backward").click()
    except:
        try:
            finall = driver.find_element(By.ID, 'captchaInput.errors').text
            print('F'+finall)
            continue
        except:
            print('G')
            status = 'N'
            note = finall
            updatetime = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            updateSQL(server,username,password,database,totb1,entitytype,status,updatetime,ID,name,note)
            print(ID,name,note,status)
            driver.find_element_by_class_name("btn.btn-primary.backward").click()
            continue


    










































