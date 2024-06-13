import io
import time
from PIL import Image
from bs4 import BeautifulSoup
from selenium import webdriver
import selenium.common.exceptions
import sys

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

server,database,username,password,totb1,entitytype = db['server'],db['database'],db['username'],db['password'],db['totb1'],db['entitytype']
url,Drvfile,imgp= wbinfo['url'],wbinfo['Drvfile'],wbinfo['imgp']
chrome_options = Options()

driver = webdriver.Chrome(Drvfile)
driver.get(url)
obs = src_obs(server,username,password,database,totb1,entitytype)

for i in foo(-1,obs-1):    
    src = dbfrom(server,username,password,database,totb1,entitytype)
    name = src[0][1]
    ID = src[0][2]
    IDN_10 = src[0][4]
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
            driver.find_element(By.ID, 'iusr').send_keys(IDN_10)
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
        driver.find_element(By.ID, 'iusr').send_keys(IDN_10)
        driver.find_element_by_name('captchaAnswer').send_keys(res)
        driver.find_element(By.ID, 'btn1').click()
        message = ''

        try:
            message=driver.switch_to_alert().text
            driver.switch_to_alert().accept()
            q = 1000
        except:
            q = 1000

    if '查無資料' in message:
        updatetime = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        login_date,login_inc,Insurance_type,status = '','查無資料','','N'
        updateSQL(server,username,password,database,totb1,entitytype,status,updatetime,ID,IDN_10,Insurance_type,login_date,login_inc)
        print(ID,IDN_10,login_date,login_inc,Insurance_type,status)
    else:
        # 不分 第一筆 or 非第一筆，統一資料處理
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # 指派查找結果之預設值
        login_date,login_inc,Insurance_type,status = '','','','N'

        # 判斷查找是否有產生結果
        if soup.find(class_ = 'formStyle02'):
            ins_info = soup.find(class_ = 'formStyle02').find_all('td')
            login_date = re.sub(r"\s+", "", ins_info[3].text)
            login_inc = re.sub(r"\s+", "", ins_info[5].text)
            updatetime = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

            if login_date != '' and '未辦理登錄' not in login_inc:
                status = 'Y'

        # 判斷查找結果列表是否有「保險種類」(Insurance_type)，8→有；6→沒有
        if len(ins_info) == 8:
            Insurance_type = re.sub(r"\s+", "", ins_info[7].text)

        # 更新結果至資料庫
        updateSQL(server,username,password,database,totb1,entitytype,status,updatetime,ID,IDN_10,Insurance_type,login_date,login_inc)
        print(ID,IDN_10,login_date,login_inc,Insurance_type,status)
        
        #計算查找筆數，今日調閱大於五千筆則停止     
        exit_o = exit_obs(server,username,password,database,totb1)
        if exit_o >= 5000:
            try:
                driver.close()
            except:
                pass
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
            try:
                driver.close()
            except:
                pass
            sys.exit()
try:
    driver.close()
except:
    pass