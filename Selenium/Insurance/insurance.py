# -*- coding: utf-8 -*-
"""
Created on Wed Apr 27 14:29:31 2022

@author: admin
"""

# -*- coding: utf-8 -*-
"""
Created on Mon Feb 14 15:47:35 2022

@author: admin
"""
import io
import time
import sys
import calendar
import requests
from PIL import Image
#from PIL import ImageGrab
from bs4 import BeautifulSoup
from selenium import webdriver
import selenium.common.exceptions
import pyscreenshot as ImageGrab
import re
import pyautogui
import datetime
import ddddocr
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.expected_conditions import alert_is_present


from config import *
from etl_func import *
from dict import *

server,database,username,password,totb,fromtb = db['server'],db['database'],db['username'],db['password'],db['totb'],db['fromtb']
url,Drvfile,imgp= wbinfo['url'],wbinfo['Drvfile'],wbinfo['imgp']
chrome_options = Options()
chrome_options.add_argument('--headless')
driver = webdriver.Chrome(Drvfile,chrome_options=chrome_options)
driver.get(url)
obs = src_obs(server,username,password,database,fromtb,totb)

mail(obs)

for i in foo(-1,obs-1):
    src = dbfrom(server,username,password,database,fromtb,totb)[0]
    today = str(datetime.datetime.now())[0:-3]
    Name = src[2]
    ID =src[1]
    print(ID)
    birthday = src[6]
    rowid = str(src[10]).replace('None','')
    print(rowid)
    bir = birthday.split('/',2)
    message = ''
    
    # 第一筆資料 retry captcha 1000次
    if i == 0:
        
        q = 0
        while q < 1000:        
            img_ele = driver.find_element_by_id('captcha')
            img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
            img.save(imgp)

            ocr = ddddocr.DdddOcr()
            with open(imgp, 'rb') as f:
                img_bytes = f.read()
            res = ocr.classification(img_bytes)

            driver.find_element(By.ID, 'iusr').clear()
            driver.find_element_by_name('captchaAnswer').clear()
            driver.find_element(By.ID, 'ibd1').clear()
            driver.find_element_by_name('BDate2').clear()
            driver.find_element_by_name('BDate3').clear()
            driver.find_element(By.ID, 'iusr').send_keys(ID)
            driver.find_element(By.ID, 'ibd1').send_keys(bir[0])
            driver.find_element_by_name('BDate2').send_keys(bir[1])
            driver.find_element_by_name('BDate3').send_keys(bir[2])
            driver.find_element_by_name('captchaAnswer').send_keys(res)
            time.sleep(1)

            driver.find_element(By.ID, 'btn1').click()
            message = ''
            time.sleep(1)

            try:
                message=driver.switch_to_alert().text
                driver.switch_to_alert().accept()

                # 驗證碼錯誤
                if '錯誤' in message:
                    driver.find_element(By.ID, 'btn3').click()
                    time.sleep(1)
                    q += 1

                # 查無資料
                else:
                    q = 1000
            except:
                q = 1000

    # 非第一筆資料
    else:
        driver.find_element(By.ID, 'iusr').clear()
        driver.find_element_by_name('captchaAnswer').clear()
        driver.find_element(By.ID, 'ibd1').clear()
        driver.find_element_by_name('BDate2').clear()
        driver.find_element_by_name('BDate3').clear()
        driver.find_element(By.ID, 'iusr').send_keys(ID)
        driver.find_element(By.ID, 'ibd1').send_keys(bir[0])
        driver.find_element_by_name('BDate2').send_keys(bir[1])
        driver.find_element_by_name('BDate3').send_keys(bir[2])
        driver.find_element_by_name('captchaAnswer').send_keys(res)
        driver.find_element(By.ID, 'btn1').click()
        message = ''

        try:
            message=driver.switch_to_alert().text
            driver.switch_to_alert().accept()
            q = 1000
        except:
            q = 1000
    print(message)
    if '查無資料' in message:
        note = 'N'
        insurance_num = ''
        docs = (Name,ID,birthday,insurance_num,today,note)
        insurance_result = insurance(docs)
        if len(rowid) > 0:
            print('A update')
            update(server,username,password,database,totb,note,ID,today,rowid,insurance_num)
            #計算查找筆數，今日調閱大於五千筆則停止     
            exit_o = exit_obs(server,username,password,database,totb)
            if exit_o >= 5000:
                driver.close()
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
                driver.close()
                sys.exit()
        else:
            toSQL(insurance_result, totb, server, database, username, password)
            #計算查找筆數，今日調閱大於五千筆則停止     
            exit_o = exit_obs(server,username,password,database,totb)
            if exit_o >= 5000:
                driver.close()
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
                driver.close()
                sys.exit()
            
        

    else:
        # 不分 第一筆 or 非第一筆，統一資料處理
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        # 指派查找結果之預設值
        login_date,login_inc,insurance_num,note = '','','','N'
        # 判斷查找是否有產生結果
        if soup.find(class_ = 'formStyle02'):
            ins_info = soup.find(class_ = 'formStyle02').find_all('td')
            try:
                insurance_num = re.sub(r"\s+", "", ins_info[1].text)
            except:
                pass
            if len(insurance_num) == 0: 
                try:
                    insurance_num = re.sub(r"\s+", "", ins_info[3].text)
                    #note = 'N'
                except:
                    #note = 'Y'
                    pass
            print(insurance_num)
            if insurance_num != '' and '未辦理登錄' not in insurance_num and '年' not in insurance_num:
                note = 'Y'
            updatetime = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

        docs = (Name,ID,birthday,insurance_num,today,note)
        insurance_result = insurance(docs)
        if len(rowid) > 0:
            update(server,username,password,database,totb,note,ID,today,rowid,insurance_num)
            #計算查找筆數，今日調閱大於五千筆則停止     
            exit_o = exit_obs(server,username,password,database,totb)
            if exit_o >= 5000:
                driver.close()
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
                driver.close()
                sys.exit()
        else:
            toSQL(insurance_result, totb, server, database, username, password)
            #計算查找筆數，今日調閱大於五千筆則停止     
            exit_o = exit_obs(server,username,password,database,totb)
            if exit_o >= 5000:
                driver.close()
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
                driver.close()
                sys.exit()
        
        print(src_obs(server, username, password, database,fromtb,totb))

        # 更新結果至資料庫
try:
    driver.close()
except:
    pass
   

