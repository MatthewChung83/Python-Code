import io
import time
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
    src = dbfrom(server,username,password,database,fromtb,totb)[0]
    #查找時間
    updatetime = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    #身份證字號
    ID = src[0]
    #姓名
    Name = str(src[1]).replace('None','')
    #rowid
    rowid = str(src[2]).replace('None','')
    #網爬姓名
    name = str(src[3]).replace('None','')
    #網爬性別
    sex = str(src[4]).replace('None','')
    #網爬公司
    company = str(src[5]).replace('None','')
    #網爬公司總類
    company_type = str(src[6]).replace('None','')
    #職員
    emp_type = str(src[7]).replace('None','')
    #受訓狀態
    tranning_type = str(src[8]).replace('None','')
    
    #辨識碼抓取
    img_ele = driver.find_element_by_id('img_validcode')
    img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
    img.save(imgp)
    ocr = ddddocr.DdddOcr()
    with open(imgp, 'rb') as f:
        img_bytes = f.read()
    res = ocr.classification(img_bytes)
    print(res)
    if str(res.isnumeric()) =='False' or len(res) !=5:
        driver.refresh()
        continue
    else:
        pass
    
    #輸入身分證字號、辨識碼
    driver.find_element_by_name('txtExamNo').clear()
    driver.find_element_by_name('txtValidate').clear()
    driver.find_element_by_name('txtExamNo').send_keys(ID)
    driver.find_element_by_name('txtValidate').send_keys(res)
    driver.find_element(By.ID, 'btnQry').click()
    time.sleep(2)
    #判斷錯誤訊息
    try:
        message=driver.switch_to_alert().text
        driver.switch_to_alert().accept()
        print(message)
    except:
        message = ''
        pass
    if '驗證碼輸入失敗!' in message:
        continue
    elif '查無相關資料!' in message:
        note = 'N'
        Name_web = ''
        sex_web = ''
        company_web = ''
        company_type_web = ''
        emp_type_web = ''
        tranning_type_web = ''
        
    else:
        soup = BeautifulSoup(driver.page_source, 'lxml')
        note = 'Y'
        Name_web = driver.find_element(By.ID, 'gvList_ctl02_lblName').text
        sex_web = driver.find_element(By.ID, 'gvList_ctl02_lblSex').text
        company_web = driver.find_element(By.ID, 'gvList_ctl02_lblCompanyName').text
        company_type_web = driver.find_element(By.ID, 'gvList_ctl02_lblCompanyType').text
        emp_type_web = driver.find_element(By.ID, 'gvList_ctl02_lblTITLE').text
        tranning_type_web = driver.find_element(By.ID, 'gvList_ctl02_lblCovid').text
        
        
    if '查無相關資料!' in message and len(rowid) > 0 and len(company_type) != 0 :
        note = 'N'
        updateSQL(server,username,password,database,totb,ID,Name_web,sex_web,company_web,company_type_web,emp_type_web,tranning_type_web,updatetime,note,rowid)
        #計算查找筆數，今日調閱大於五千筆則停止     
        exit_o = exit_obs(server,username,password,database,totb)
        if exit_o >= 5000:
            driver.quit()
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
            driver.quit()
            sys.exit()

    elif '查無相關資料!' not in message and len(rowid) > 0 and name == Name_web and sex == sex_web and company == company_web and company_type == company_type_web and emp_type == emp_type_web and tranning_type == tranning_type_web:
        note = 'Y'
        updateSQL(server,username,password,database,totb,ID,Name_web,sex_web,company_web,company_type_web,emp_type_web,tranning_type_web,updatetime,note,rowid)
        #計算查找筆數，今日調閱大於五千筆則停止     
        exit_o = exit_obs(server,username,password,database,totb)
        if exit_o >= 5000:
            driver.quit()
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
            driver.quit()
            sys.exit()
    else:
        docs = (ID,Name,Name_web,sex_web,company_web,company_type_web,emp_type_web,tranning_type_web,updatetime,note)
        travelagency_result = travelagency(docs)
        toSQL(travelagency_result, totb, server, database, username, password)
        #計算查找筆數，今日調閱大於五千筆則停止     
        exit_o = exit_obs(server,username,password,database,totb)
        if exit_o >= 5000:
            driver.quit()
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
            driver.quit()
            sys.exit()
try:
    driver.close()
except:
    pass