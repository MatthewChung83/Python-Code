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
check = check(server, username, password, database)
print(check)
driver = webdriver.Chrome(Drvfile)

driver.get(url)


obs = src_obs(server,username,password,database,fromtb,totb,check)
num(obs,server,username,password,database,fromtb,totb,check)
mail(obs)
for i in foo(-1,obs*2):
    try:
        obs = src_obs(server,username,password,database,fromtb,totb,check)
        src = dbfrom(server,username,password,database,fromtb,totb,check)[0]
        today = str(datetime.datetime.now())[0:-3]
        Name = src[2]
        ID =src[1]
        print(ID)
        birthday = src[6]
        rowid = str(src[9]).replace('None','')
        print(rowid)
        bir = birthday.split('/',2)
        driver.find_element(By.ID, 'iusr').clear()
        driver.find_element(By.ID, 'ibd1').clear()
        driver.find_element_by_name('BDate2').clear()
        driver.find_element_by_name('BDate3').clear()
        driver.find_element(By.ID, 'iusr').send_keys(ID)
        driver.find_element(By.ID, 'ibd1').send_keys(bir[0])
        driver.find_element_by_name('BDate2').send_keys(bir[1])
        driver.find_element_by_name('BDate3').send_keys(bir[2])
        img_ele = driver.find_element_by_id('captcha')
        img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
        img.save(imgp)
        
        ocr = ddddocr.DdddOcr()
        with open(imgp, 'rb') as f:
            img_bytes = f.read()
        res = ocr.classification(img_bytes)
        print(res)
        driver.find_element_by_name('captchaAnswer').send_keys(res)
        driver.find_element(By.ID, 'btn1').click()
        #message = []
        #while message=='驗證碼正確' or '查無資料' in message :
        try:
            message=driver.switch_to_alert().text
            #print(message)
        except:
            message='驗證碼正確'
        
        if '查無資料' in message :
            note = 'N'
            insurance_num = ''
            docs = (Name,ID,birthday,insurance_num,today,note)
            insurance_result = insurance(docs)
            if len(rowid) > 0:
                update(server,username,password,database,totb,note,ID,today,rowid)
            else:
                toSQL(insurance_result, totb, server, database, username, password)
            
            print(src_obs(server, username, password, database,fromtb,totb,check))
            pyautogui.press('enter')
            
            for i in foo(-1,obs-1):
                src = dbfrom(server,username,password,database,fromtb,totb,check)[0]
                rowid = i
                Name = src[2]
                ID = src[1]
                birthday = src[6]
                bir = str(birthday).split('/',2)
                driver.find_element(By.ID, 'iusr').clear()
                driver.find_element(By.ID, 'ibd1').clear()
                driver.find_element_by_name('BDate2').clear()
                driver.find_element_by_name('BDate3').clear()
                        
                driver.find_element(By.ID, 'iusr').send_keys(ID)
                driver.find_element(By.ID, 'ibd1').send_keys(bir[0])
                driver.find_element_by_name('BDate2').send_keys(bir[1])
                driver.find_element_by_name('BDate3').send_keys(bir[2])
                driver.find_element(By.ID, 'btn1').click()
                try:
                    message=driver.switch_to_alert().text
                    #print(message)
                except:
                    message='驗證碼正確'
                if '查無資料' in message :
                    note = 'N'
                    insurance_num = ''
                    docs = (Name,ID,birthday,insurance_num,today,note)
                    insurance_result = insurance(docs)
                    if len(rowid) > 0:
                        update(server,username,password,database,totb,note,ID,today,rowid)
                    else:
                        toSQL(insurance_result, totb, server, database, username, password)
                    #print(src_obs(server, username, password, database,fromtb,totb))
                    pyautogui.press('enter')
                    obs = src_obs(server,username,password,database,fromtb,totb,check)
                    num(obs,server,username,password,database,fromtb,totb,check)
                    pass
                elif '未辦理登錄' in message :
                    note = 'N'
                    insurance_num = ''
                    docs = (Name,ID,birthday,insurance_num,today,note)
                    insurance_result = insurance(docs)
                    if len(rowid) > 0:
                        update(server,username,password,database,totb,note,ID,today,rowid)
                    else:
                        toSQL(insurance_result, totb, server, database, username, password)
                    #print(src_obs(server, username, password, databas,fromtb,totbe))
                    pyautogui.press('enter')
                    obs = src_obs(server,username,password,database,fromtb,totb,check)
                    num(obs,server,username,password,database,fromtb,totb,check)
                    pass
                else:
                    try:
                        insurance_num = driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/span/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]').text
                    except:
                        insurance_num = driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/span/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[2]').text                    
                    note = 'Y'
                    docs = (Name,ID,birthday,insurance_num,today,note)
                    insurance_result = insurance(docs)
                    print(insurance_result)
                    if len(rowid) > 0:
                        update(server,username,password,database,totb,note,ID,today,rowid)
                    else:
                        toSQL(insurance_result, totb, server, database, username, password)
                    print(src_obs(server, username, password, database,fromtb,totb,check))
                    obs = src_obs(server,username,password,database,fromtb,totb,check)
                    num(obs,server,username,password,database,fromtb,totb,check)
                    pass
        elif '驗證碼正確' in message :
            try:
                insurance_num = driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/span/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]').text
                note = 'N'
                docs = (Name,ID,birthday,insurance_num,today,note)
                insurance_result = insurance(docs)
                print(insurance_result)
                if len(rowid) > 0:
                    update(server,username,password,database,totb,note,ID,today,rowid)
                else:
                    toSQL(insurance_result, totb, server, database, username, password)
                print(src_obs(server, username, password, database,fromtb,totb,check))
                pass
            except:
                insurance_num = driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/span/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[2]').text                    
                note = 'Y'
                docs = (Name,ID,birthday,insurance_num,today,note)
                insurance_result = insurance(docs)
                print(insurance_result)
                if len(rowid) > 0:
                    update(server,username,password,database,totb,note,ID,today,rowid)
                else:
                    toSQL(insurance_result, totb, server, database, username, password)
                print(src_obs(server, username, password, database,fromtb,totb,check))
                pass
            for i in foo(-1,obs-1):
                src = dbfrom(server,username,password,database,fromtb,totb,check)[0]
                rowid = i
                Name = src[2]
                ID = src[1]
                birthday = src[6]
                bir = str(birthday).split('/',2)
                driver.find_element(By.ID, 'iusr').clear()
                driver.find_element(By.ID, 'ibd1').clear()
                driver.find_element_by_name('BDate2').clear()
                driver.find_element_by_name('BDate3').clear()
                        
                driver.find_element(By.ID, 'iusr').send_keys(ID)
                driver.find_element(By.ID, 'ibd1').send_keys(bir[0])
                driver.find_element_by_name('BDate2').send_keys(bir[1])
                driver.find_element_by_name('BDate3').send_keys(bir[2])
                driver.find_element(By.ID, 'btn1').click()
                try:
                    message=driver.switch_to_alert().text
                    #print(message)
                except:
                    message='驗證碼正確'
                if '查無資料' in message :
                    note = 'N'
                    insurance_num = ''
                    docs = (Name,ID,birthday,insurance_num,today,note)
                    insurance_result = insurance(docs)
                    if len(rowid) > 0:
                        update(server,username,password,database,totb,note,ID,today,rowid)
                    else:
                        toSQL(insurance_result, totb, server, database, username, password)
                    #print(src_obs(server, username, password, databas,fromtb,totbe))
                    pyautogui.press('enter')
                    obs = src_obs(server,username,password,database,fromtb,totb,check)
                    num(obs,server,username,password,database,fromtb,totb,check)
                    pass
                elif '未辦理登錄' in message :
                    note = 'N'
                    insurance_num = ''
                    docs = (Name,ID,birthday,insurance_num,today,note)
                    insurance_result = insurance(docs)
                    if len(rowid) > 0:
                        update(server,username,password,database,totb,note,ID,today,rowid)
                    else:
                        toSQL(insurance_result, totb, server, database, username, password)
                    #print(src_obs(server, username, password, databas,fromtb,totbe))
                    pyautogui.press('enter')
                    obs = src_obs(server,username,password,database,fromtb,totb,check)
                    num(obs,server,username,password,database,fromtb,totb,check)
                    pass
                else:
                    try:
                        insurance_num = driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/span/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]').text
                    except:
                        insurance_num = driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/span/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[2]').text                    
                    note = 'Y'
                    docs = (Name,ID,birthday,insurance_num,today,note)
                    insurance_result = insurance(docs)
                    print(insurance_result)
                    if len(rowid) > 0:
                        update(server,username,password,database,totb,note,ID,today,rowid)
                    else:
                        toSQL(insurance_result, totb, server, database, username, password)
                    #print(src_obs(server, username, password, database,fromtb,totb))
                    obs = src_obs(server,username,password,database,fromtb,totb,check)
                    num(obs,server,username,password,database,fromtb,totb,check)
                    pass
        elif '驗證碼錯誤' in message :
            pyautogui.press('enter')
            driver.refresh()
            continue 
        elif '出生日期格式錯誤' in message :
            note = 'N'
            insurance_num = ''
            docs = (Name,ID,birthday,insurance_num,today,note)
            insurance_result = insurance(docs)
            if len(rowid) > 0:
                update(server,username,password,database,totb,note,ID,today,rowid)
            else:
                toSQL(insurance_result, totb, server, database, username, password)
                
                
                
                
            #print(src_obs(server, username, password, database,fromtb,totb))
            pyautogui.press('enter')
            continue
        else :
            pyautogui.press('enter')
            driver.refresh()
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
    where Engine = 'insurance' and type = '{check}'
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    errormail(obs)
    driver.quit()
elif obs ==0 :
    
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update UCS_Source
    set status = 'Done'
    where Engine = 'insurance' and type = '{check}'
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()    
    errormail(obs)
    driver.quit()
