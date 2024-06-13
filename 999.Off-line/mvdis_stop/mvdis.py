# -*- coding: utf-8 -*-
"""
Created on Tue Feb 15 14:07:33 2022

@author: admin
"""
import io
import time
import datetime
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

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.expected_conditions import alert_is_present


from config import *
from etl_func import *
from dict import *


server,database,username,password,fromtb,totb = db['server'],db['database'],db['username'],db['password'],db['fromtb'],db['totb']
url,Drvfile,imgp= wbinfo['url'],wbinfo['Drvfile'],wbinfo['imgp']
check = check(server, username, password, database)
driver = webdriver.Chrome(Drvfile)
driver.get(url)
obs = src_obs(server,username,password,database,fromtb,totb,check)
print(obs)
mail(obs)

for i in foo(-1,obs*2):
    try:
        obs = src_obs(server,username,password,database,fromtb,totb,check)
        src = dbfrom(server,username,password,database,fromtb,totb,check)[0]
        num(obs,server,username,password,database,fromtb,totb,check)
        today = datetime.datetime.now()
        #rowid = i
        Name = src[2]
        ID = src[1]
        birthday = src[6]
        a=birthday.split('-',1)
        rowid = src[9]
        print(Name)
        print(ID)
        print(birthday)
        if len(ID) == 10:
            if len(a[0]) == 3:
                bir = birthday.replace('-','')
            else : 
                bir = '0'+birthday.replace('-','')
            
            driver.find_element(By.ID, 'idNo').clear()
            driver.find_element(By.ID, 'birthday').clear()
            driver.find_element(By.ID, 'idNo').send_keys(ID)
            driver.find_element(By.ID, 'birthday').send_keys(bir)
            windows = driver.window_handles
            driver.switch_to.window(windows[-1])
            driver.maximize_window()
            time.sleep(1)
            pyautogui.press('enter')
            
            img_ele = driver.find_element_by_id('pickimg1')
            img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
            img.save(imgp)
            import ddddocr
            ocr = ddddocr.DdddOcr()
            with open(imgp, 'rb') as f:
                img_bytes = f.read()
            res = ocr.classification(img_bytes)
            print(res)
            driver.find_element_by_name('validateStr').send_keys(res)
            driver.find_element(By.ID, 'submit_btn').click()
            #message = driver.find_elements_by_xpath("html/body/table/tbody/tr[2]/td/div[2]/div/table/tbody")
            try:
                message=driver.find_element_by_id('validateStr.errors').text
                print(message)
            except:
                message=driver.find_element(By.ID, 'headerMessage').text
                print(message)
                if '驗證碼輸入錯誤' in message :
                    driver.refresh()
                    pass
                elif '請確認您輸入的證號及生日是否正確。' in message :
                    note = 'N'
                    update_date = today
                    docs = (Name,ID,birthday,update_date,note)
                    dead_result = dead(docs)
                    toSQL(dead_result, totb, server, database, username, password)
                    update(server,username,password,database,totb,note,ID,today)
                    #print(src_obs(server, username, password, database))
                    
                    pass
                else:
                    message1 = driver.find_elements_by_xpath("html/body/table/tbody/tr[2]/td/div[2]/div/table/tbody/tr/td")
                    message_F = []
                    for i in message1:
                        #message_F = i.text
                        
                        message_F.append(i.text)
                        print(message_F)
                        
                    if '免換照' in message_F:
                        note = 'N'
                        update_date = today
                        docs = (Name,ID,birthday,update_date,note)
                        dead_result = dead(docs)
                        toSQL(dead_result, totb, server, database, username, password)
                        update(server,username,password,database,totb,note,ID,today)
                        #print(src_obs(server, username, password, database))
                        driver.find_element_by_class_name('std_btn')
                        driver.refresh()
                        pass
                    elif '此車非活車' in message_F:
                        note = 'Y'
                        update_date = today
                        docs = (Name,ID,birthday,update_date,note)
                        dead_result = dead(docs)
                        toSQL(dead_result, totb, server, database, username, password)
                        update(server,username,password,database,totb,note,ID,today)
                        #print(src_obs(server, username, password, database,fromtb,totb))
                        driver.find_element_by_class_name('std_btn')
                        driver.refresh()
                        pass
                        
    
                    else  :
                        note = 'N'
                        update_date = today
                        docs = (Name,ID,birthday,update_date,note)
                        dead_result = dead(docs)
                        toSQL(dead_result, totb, server, database, username, password)
                        update(server,username,password,database,totb,note,ID,today)
                        #print(src_obs(server, username, password, database,fromtb,totb))
                        driver.find_element_by_class_name('std_btn')
                        driver.refresh()
                        pass
        else:
            note = 'N'
            update_date = today
            docs = (Name,ID,birthday,update_date,note)
            dead_result = dead(docs)
            toSQL(dead_result, totb, server, database, username, password)
            update(server,username,password,database,totb,note,ID,today)
            #print(src_obs(server, username, password, database,fromtb,totb))
            pass
    except:
        pass
    
obs = src_obs(server,username,password,database,fromtb,totb,check)
if obs > 0 :
    
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update UCS_Source
    set F_OBS = '{obs}'
    where Engine = 'dead' and type = '{check}'
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    errormail(obs)
    
elif obs ==0 :
    
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update UCS_Source
    set status = 'Done'
    where Engine = 'dead' and type = '{check}'
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()    
    errormail(obs)
    driver.quit()