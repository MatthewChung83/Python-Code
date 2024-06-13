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
    src = dbfrom(server,username,password,database,fromtb,totb)[0]
    #查找時間
    updatetime = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    #身份證字號
    ID = src[0]
    #姓名
    Name = str(src[1]).replace('None','').replace('','')
    #rowid
    rowid = str(src[2]).replace('None','')
    #網爬註記
    notation = str(src[3]).replace('None','')
    #網爬建築師-姓名
    architect_Name = str(src[4]).replace('None','')
    #網爬建築師證號
    license_num = str(src[5]).replace('None','')
    #網爬開業證號
    certificate_num = str(src[6]).replace('None','')
    #辦公室
    office_name = str(src[7]).replace('None','')
    
    #首筆資料
    if i  == 0 :
        driver.find_element_by_class_name('btn-bar.pc-hide').click()
        time.sleep(1)
        driver.find_element_by_class_name('btn-more.pc-hide').click()
        time.sleep(1)
        driver.find_element_by_class_name('menu-item-href').click()
        time.sleep(1)
    else:
        pass
    
    #辨識碼抓取
    img_ele = driver.find_element_by_id('vadimg')
    img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
    img.save(imgp)
    ocr = ddddocr.DdddOcr()
    with open(imgp, 'rb') as f:
        img_bytes = f.read()
    res = ocr.classification(img_bytes)
    print(res.upper())


    #輸入身分證字號、姓名、辨識碼
    driver.find_element_by_name('Qry_ID_NO').clear()
    driver.find_element_by_name('Qry_NAME').clear()
    driver.find_element_by_name('Qry_ID_NO').send_keys(ID)
    driver.find_element_by_name('Qry_NAME').send_keys(Name)
    driver.find_element_by_name('Qry_imageCodetxt').send_keys(res.upper())
    driver.find_element_by_name('QueryParamButton_executeQuery').click()
    time.sleep(1)
    #判斷錯誤訊息
    message = driver.find_element_by_id('warning_txt').text
    if '驗證碼' in message or len(message)>0:
        driver.refresh()
        continue
    else:
        pass
    soup = BeautifulSoup(driver.page_source, 'lxml')
    try:
        check = soup.find(id = 'body_table').find_all('td')
    except:
        check = ''
    if len(check) == 0 :
        note = 'N'
        notation_web = ''
        architect_Name_web = ''
        license_num_web = ''
        certificate_num_web = ''
        office_name_web = ''

    else :
        note = 'Y'
        notation_web = soup.find(id = 'body_table').find_all('td')[0].text
        architect_Name_web = soup.find(id = 'body_table').find_all('td')[1].text
        license_num_web = soup.find(id = 'body_table').find_all('td')[2].text
        certificate_num_web = soup.find(id = 'body_table').find_all('td')[3].text
        office_name_web = soup.find(id = 'body_table').find_all('td')[4].text

        
        
    if  len(check) == 0 and len(rowid) > 0 and len(architect_Name_web) == 0 and len(license_num_web) == 0 and len(certificate_num_web) ==0  :
        updateSQL(server,username,password,database,totb,ID,notation_web,architect_Name_web,license_num_web,certificate_num_web,office_name_web,updatetime,note,rowid)
        #計算查找筆數，今日調閱大於五千筆則停止     
        exit_o = exit_obs(server,username,password,database,totb)
        if exit_o >= 10000:
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
            sys.exit()

    elif  len(check) != 0 and len(rowid) > 0 and architect_Name_web == architect_Name and license_num_web == license_num and certificate_num_web == certificate_num and office_name_web == office_name :
        updateSQL(server,username,password,database,totb,ID,notation_web,architect_Name_web,license_num_web,certificate_num_web,office_name_web,updatetime,note,rowid)
        #計算查找筆數，今日調閱大於五千筆則停止     
        exit_o = exit_obs(server,username,password,database,totb)
        if exit_o >= 10000:
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
            sys.exit()
    else:
        docs = (ID,Name,notation_web,architect_Name_web,license_num_web,certificate_num_web,office_name_web,updatetime,note)
        cpami_result = cpami(docs)
        toSQL(cpami_result, totb, server, database, username, password)
        #計算查找筆數，今日調閱大於五千筆則停止     
        exit_o = exit_obs(server,username,password,database,totb)
        if exit_o >= 10000:
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
            sys.exit()
try:
    driver.close()
except:
    pass