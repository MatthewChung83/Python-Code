# -*- coding: utf-8 -*-
"""
Created on Wed Jan 19 15:08:27 2022

@author: admin
"""


import io
import time
import os
import sys
from PIL import Image
from bs4 import BeautifulSoup
from selenium import webdriver
import selenium.common.exceptions

import re
import datetime
import ddddocr
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.expected_conditions import alert_is_present
import time

from config import *
from etl_func import *
from dict import *


server,database,username,password,fromtb,fromtb1,totb = db['server'],db['database'],db['username'],db['password'],db['fromtb'],db['fromtb1'],db['totb']
url,Drvfile,imgp= wbinfo['url'],wbinfo['Drvfile'],wbinfo['imgp']

#obs = src_obs(server,username,password,database,fromtb,fromtb1,totb)
#mail(obs)


    
#dbfrom(server,username,password,database,fromtb,fromtb1,totb)
driver = webdriver.Chrome(Drvfile)
driver.get(url)
obs_U = src_obs_U(server,username,password,database,totb)   
for i in foo(-1,obs_U-1):
    today = str(datetime.datetime.now())[0:-3]
    a,b = 1,1
    try:
        while a < 100:
            message = ''
            #辨識碼抓取
            while b < 100:
                img_ele = driver.find_element_by_xpath('//*[@id="etwMainContent"]/div[2]/div/div[2]/jhi-main/etw113w1-name-query/form/div[1]/div[2]/div[4]/div/div/etw-captcha/div/img')
                img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
                img.save(imgp)
                ocr = ddddocr.DdddOcr()
                with open(imgp, 'rb') as f:
                    img_bytes = f.read()
                res = ocr.classification(img_bytes)
                print(res)
                if len(res)<7:
                    b = 100
                else:
                    driver.find_element_by_xpath('//*[@id="etwAsideNav"]/div[2]/ul/li[2]/a').click()
                    time.sleep(2)
                    b=b+1
            src_U = dbfrom_U(server,username,password,database,totb)[0]  
            query_company = src_U[0]
            print(query_company)
            #清除
            driver.find_element_by_xpath('//*[@id="name"]').clear()
            driver.find_element_by_xpath('//*[@id="captchaText"]').clear()
            #輸入值
            driver.find_element_by_xpath('//*[@id="name"]').send_keys(query_company)
            driver.find_element_by_xpath('//*[@id="captchaText"]').send_keys(res)
            
            #點選查詢
            driver.find_element_by_xpath('//*[@id="etwMainContent"]/div[2]/div/div[2]/jhi-main/etw113w1-name-query/form/div[2]/div[2]/button').click()
            
            #sleep
            time.sleep(10)
            try:
                message = driver.find_element_by_xpath('/html/body/ngb-modal-window/div/div/jhi-dialog/div/div[2]').text
                driver.find_element_by_xpath('/html/body/ngb-modal-window/div/div/jhi-dialog/div/div[3]/div/button').click()
            except:
                pass
            if '驗證碼錯誤' in message:
                driver.find_element_by_xpath('//*[@id="etwAsideNav"]/div[2]/ul/li[2]/a').click()
                a = a+1
                b = 1
            else :
                a = 100
        
        if '查無資料' in message :
            Thrid_Company = ''
            Thrid_Name = ''
            Thrid_Num = ''
            Business_address = ''
            Business_status = ''
            Capital_amount = ''
            Tissue_type = ''
            Date_of_establishment = ''
            Register_business_items = ''
            Note,casei,Legal_status,Court,status= '','','','','N'
            update(server,username,password,database,totb,query_company,Thrid_Company,Thrid_Name,Thrid_Num,Business_address,Business_status,Capital_amount,Tissue_type,Date_of_establishment,Register_business_items,today,status)
        #elif '驗證碼錯誤' in message :
        #    obs_U = src_obs_U(server,username,password,database,totb)     
        #    time.sleep(1)
        #    driver.find_element_by_xpath('//*[@id="etwAsideNav"]/div[2]/ul/li[2]/a').click()
        #    continue
        else:
            
            #點選連結
            driver.find_element_by_xpath('//*[@id="etwMainContent"]/div[2]/div/div[2]/jhi-main/etw113w1-name-query-result/div[1]/div[2]/table/tbody/tr/td[2]').click()
            
            time.sleep(3)
            #抓取頁面資料
            soup = BeautifulSoup(driver.page_source, 'lxml')
            
            #Thrid_Company
            try:
                Thrid_Company = soup.find_all('ul')[4].find_all('li')[3].find_all('div')[1].text
            except:
                Thrid_Company = ''
            
            #Thrid_Name
            try:
                Thrid_Name = soup.find_all('ul')[4].find_all('li')[2].find_all('div')[1].text
            except:
                Thrid_Name = ''
            #Thrid_Num
            try:
                Thrid_Num = soup.find_all('ul')[4].find_all('li')[0].find_all('div')[1].text
            except:
                Thrid_Num = ''
                
            #Business_address
            try:
                Business_address = soup.find_all('ul')[4].find_all('li')[4].find_all('div')[1].text
            except:
                Business_address = ''
            #Business_status
            try:
                Business_status = soup.find_all('ul')[4].find_all('li')[1].find_all('div')[1].text
            except:
                Business_status = ''
                
            #Capital_amount
            try:
                Capital_amount = soup.find_all('ul')[4].find_all('li')[5].find_all('div')[1].text
            except:
                Capital_amount = ''
            #Tissue_type
            try:
                Tissue_type = soup.find_all('ul')[4].find_all('li')[6].find_all('div')[1].text
            except:
                Tissue_type = ''
            #Date_of_establishment
            try:
                Date_of_establishment = soup.find_all('ul')[4].find_all('li')[7].find_all('div')[1].text
            except:
                Date_of_establishment = ''
            #Register_business_items
            try:
                Register_business_items = soup.find_all('ul')[4].find_all('li')[8].find_all('div')[1].text
            except:
                Register_business_items = ''
            #預設空值
            Note,casei,Legal_status,Court,status = '','','','','Y'
        
            update(server,username,password,database,totb,query_company,Thrid_Company,Thrid_Name,Thrid_Num,Business_address,Business_status,Capital_amount,Tissue_type,Date_of_establishment,Register_business_items,today,status)
            
            driver.find_element_by_xpath('//*[@id="etwAsideNav"]/div[2]/ul/li[2]/a').click()
    except:
        driver.close()
        os.system(rf"C:\Py_Project\project\etwmain_03\etwmain_03.bat")
        
        pass



driver.close()





















