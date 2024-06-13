import io
import time
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


from config import *
from etl_func import *
from dict import *

server,database,username,password,totb,fromtb = db['server'],db['database'],db['username'],db['password'],db['totb'],db['fromtb']
url,Drvfile,imgp= wbinfo['url'],wbinfo['Drvfile'],wbinfo['imgp']
chrome_options = Options()

driver = webdriver.Chrome(Drvfile)
driver.get(url)
obs = src_obs(server,username,password,database,fromtb,totb)

for i in foo(-1,obs-1):   
    birthday_message = ''
    id_message = ''
    src = dbfrom(server,username,password,database,fromtb,totb)[0]
    #查找時間
    updatetime = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    #身份證字號
    if len(src[0]) == 10 and str(src[0][0].isnumeric()) =='False' and str(src[0][1].isnumeric()) =='True' :
        ID = src[0]
    else:
        ID = src[0]
        id_message = '身分證不符'
    #債務人姓名
    name = src[1]
    #生日
    if len(src[2].split('/')[0]) == 2 :
        birthday = '0'+str(src[2].replace('/',''))
    elif len(src[2].split('/')[0]) == 3:
        birthday = str(src[2].replace('/',''))
    else:
        birthday = '0'+str(src[2].replace('/',''))
        birthday_message = '生日不符'
    #rowid
    rowid = str(src[3]).replace('None','')
    #保險狀態(核保)
    Insurance_type = str(src[4]).replace('None','')
    #核保保險單號
    Insurance_num = str(src[5]).replace('None','')
    #保險狀態(理賠)
    Insurance_query_type = str(src[6]).replace('None','')
    #理賠保險單號
    Insurance_query_num = str(src[7]).replace('None','') 
    
    if '生日不符' not in birthday_message and '身分證不符' not in id_message:
        
        #頁籤切換(核保)
        driver.find_element_by_link_text("核保").click()
        #辨識碼抓取
        img_ele = driver.find_element_by_id('ValidatePic')
        img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
        img.save(imgp)
        ocr = ddddocr.DdddOcr()
        with open(imgp, 'rb') as f:
            img_bytes = f.read()
        res = ocr.classification(img_bytes)
        
        #輸入身分證字號、生日、辨識碼
        driver.find_element_by_name('AssuredId_Query').send_keys(ID)
        driver.find_element_by_name('AssuredBirth_Query').send_keys(birthday)
        driver.find_element_by_name('verifiyCode').send_keys(res)
        #查詢
        driver.find_element(By.ID, 'btnQuery').click()
        try:
            driver.find_element_by_xpath('/html/body/div[4]/div/div[3]/button[1]').click()
            message = '查無資料'
            driver.refresh()
        except:
            message = ''
            pass
        
        if '查無資料' not in message:
            note = 'Y'
            #當前頁面內容(核保)
            soup = BeautifulSoup(driver.page_source, 'lxml')
            rows = soup.findAll('tr')
            for row in rows:
                all_tds = row.find_all('td') 
                if len(all_tds) > 0:
                    Insurance_type_web = '核保'
                    ID_web = all_tds[0].text
                    Insurance_num_web = all_tds[1].text
                    Insurance_status_web = all_tds[2].text
            #回首頁
            driver.find_element(By.ID, 'btnBack').click()
            #頁籤切換(理賠)
            driver.find_element_by_link_text("理賠").click()
            #辨識碼抓取
            img_ele = driver.find_element_by_id('ValidatePic')
            img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
            img.save(imgp)
            ocr = ddddocr.DdddOcr()
            with open(imgp, 'rb') as f:
                img_bytes = f.read()
            res = ocr.classification(img_bytes)
            
            #輸入身分證字號、生日、辨識碼
            driver.find_element_by_name('AssuredId_Status').send_keys(ID)
            driver.find_element_by_name('AssuredBirth_Status').send_keys(birthday)
            driver.find_element_by_name('verifiyCode').send_keys(res)
            #查詢
            driver.find_element(By.ID, 'btnClaim').click()
            try:
                driver.find_element_by_xpath('/html/body/div[4]/div/div[3]/button[1]').click()
                Insurance_query_type_web,ID_query_web,Insurance_query_num_web,Insurance_query_status_web='','','',''
                driver.refresh()
            except:
                #當前頁面內容(理賠)
                soup_Status = BeautifulSoup(driver.page_source, 'lxml')
                rows = soup_Status.findAll('tr')
                for row in rows:
                    all_tds = row.find_all('td') 
                    if len(all_tds) > 0:
                        Insurance_query_type_web ='理賠'
                        ID_query_web = all_tds[0].text
                        Insurance_query_num_web = all_tds[1].text
                        Insurance_query_status_web = all_tds[2].text
                driver.find_element(By.ID, 'btnBack').click()
                driver.refresh()
                pass
            
        else:
            note = 'N'
            Insurance_type_web,ID_web,Insurance_num_web,Insurance_status_web = '','','',''
            Insurance_query_type_web,ID_query_web,Insurance_query_num_web,Insurance_query_status_web = '','','',''
            pass
    else:
        message = ''
        note = 'N'
        Insurance_type_web,ID_web,Insurance_num_web,Insurance_status_web = '','','',''
        Insurance_query_type_web,ID_query_web,Insurance_query_num_web,Insurance_query_status_web = '','','',''
        docs = (name,ID,birthday,Insurance_type_web,Insurance_num_web,Insurance_status_web,Insurance_query_type_web,Insurance_query_num_web,Insurance_query_status_web,updatetime,note)
        tmnewa_result = tmnewa(docs)
        toSQL(tmnewa_result, totb, server, database, username, password)
        continue
    if '查無資料' in message and len(rowid) > 0 and len(Insurance_num) == 0 :
        note = 'N'
        updateSQL(server,username,password,database,totb,note,ID,updatetime,rowid,Insurance_type_web,Insurance_num_web,Insurance_status_web,Insurance_query_type_web,ID_query_web,Insurance_query_num_web,Insurance_query_status_web)

    elif '查無資料' not in message and len(rowid) > 0 and Insurance_num == Insurance_num_web and Insurance_query_num == Insurance_query_num_web and note == 'Y':
        note = 'Y'
        updateSQL(server,username,password,database,totb,note,ID,updatetime,rowid,Insurance_type_web,Insurance_num_web,Insurance_status_web,Insurance_query_type_web,ID_query_web,Insurance_query_num_web,Insurance_query_status_web)
    else:
        
        docs = (name,ID,birthday,Insurance_type_web,Insurance_num_web,Insurance_status_web,Insurance_query_type_web,Insurance_query_num_web,Insurance_query_status_web,updatetime,note)
        tmnewa_result = tmnewa(docs)
        toSQL(tmnewa_result, totb, server, database, username, password)
try:
    driver.close()
except:
    pass