# Browser Driver
import re
import gc
import csv
import time
import pymssql
import selenium
import numpy as np
import pandas as pd
from PIL import Image
from bs4 import BeautifulSoup
from selenium import webdriver

from datetime import datetime as dt
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import io


from config import *
from etl_func import *
from OCR_Mod import *

DriverPath = r'C:\Py_Project\env\chromedriver_win32\chromedriver.exe'
Url = 'https://svc.tax.nat.gov.tw/svc/IbxPaidQuery.jsp'
chrome_options = Options()
chrome_options.add_argument('--headless')
driver = webdriver.Chrome(DriverPath,chrome_options=chrome_options)
driver.get(Url)

#clf = models.load_model(r'C:\Py_Project\project\Tax_Refund\model\tax_best_model.h5')
#clf = models.load_model(r'C:\Py_Project\project\Tax_Refund\model\etax_nat_query_model.h5')

server = db['server']
database = db['database']
username = db['username']
password = db['password']
entitytype = db['entitytype']
fromtb = db['fromtb']
totb = db['totb']

Apurl,imgf1,imgf2,imgp1,imgp2 = APinfo['Apurl'],APinfo['imgf1'],APinfo['imgf2'],APinfo['imgp1'],APinfo['imgp2']

obs = src_obs(server,username,password,database,fromtb,totb,entitytype)
try:
    for i in foo(-1,obs-1):
        src = dbfrom(server,username,password,database,fromtb,totb,entitytype)
        psid = src[0][0]
        pid = src[0][1]
        birthyear_tw = src[0][10]
        
        driver.find_element_by_name('idn').send_keys(pid)
        time.sleep(0.5)
        driver.find_element_by_name('bornYr').send_keys(birthyear_tw)
        
        q = 0
        while q < 30:
            img_ele = driver.find_element_by_id('validateCode')
            img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
            
            img_ele = driver.find_element_by_id('validateCode')
            img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
            img_1 = img.crop((3,2,130,35))
            img_1.save(imgp1)
            img_2 = img.crop((25,2,150,35))
            img_2.save(imgp2)
            
            captchaocr1 = CAPTCHAOCR(Apurl,imgf1,imgp1).replace('"','')
            captchaocr2 = CAPTCHAOCR(Apurl,imgf2,imgp2).replace('"','')
            
            l1 = len(captchaocr1)
            l2 = len(captchaocr2)
            
            while (l1 != 5 or l2 != 5):
                driver.find_element_by_id('validateCode_text').click()
                img_ele = driver.find_element_by_id('validateCode')
                img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
                
                img_ele = driver.find_element_by_id('validateCode')
                img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
                img_1 = img.crop((3,2,130,35))
                img_1.save(imgp1)
                img_2 = img.crop((25,2,150,35))
                img_2.save(imgp2)
                
                captchaocr1 = CAPTCHAOCR(Apurl,imgf1,imgp1).replace('"','')
                captchaocr2 = CAPTCHAOCR(Apurl,imgf2,imgp2).replace('"','')
                
                l1 = len(captchaocr1)
                l2 = len(captchaocr2)
                
            captchaocr = captchaocr1[0:5] + captchaocr2[-1]
            
            driver.find_element_by_id('inputCode').send_keys(captchaocr)
            driver.find_element_by_id('button').click()
            
            try:
                WebDriverWait(driver, 3).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                print(alert.text)
                time.sleep(0.5)
 
                
                if '身分證不正確' in alert.text:
                    info = alert.text
                    print(info)
                    status = 'done'
                    INQCode,Taxreturnchannel,Taxreturnplat,taxreturndate,taxOffice,taxOfficeAddr,taxOfficePhone = '','','','','','',''
                    updatesql(server,username,password,database,status,info,INQCode,Taxreturnchannel,Taxreturnplat,taxreturndate,taxOffice,taxOfficeAddr,taxOfficePhone,entitytype,psid,pid)
                    driver.find_element_by_xpath('//*[@id="cmxform"]/p[3]/input[1]').click()
                else :
     
                    pass
                alert.accept()
  
                time.sleep(0.5)

            
            except TimeoutException:
                break

            
            
        soup = BeautifulSoup(driver.page_source,'html.parser')
        if '您已完成' in soup.find('fieldset').find('p').text or '稅額試算通知書檔案編號' in soup.find('fieldset').find('p').text:
            info = soup.find('fieldset').find('p').text
            status = 'done'
            output = []
            for t in soup.find('fieldset').find_all('p'):
                output.append(t.text.replace('\n','').replace('\t','').replace('\xa0',''))
                
            INQCode,Taxreturnchannel,Taxreturnplat,taxreturndate,taxOffice,taxOfficeAddr,taxOfficePhone='','','','','','',''
            for v in output:
                if '稅額試算通知書檔案編號' in v or '檔案編號' in v:
                    INQCode = v.replace('稅額試算通知書檔案編號','').replace('檔案編號','').replace('： ','')               
                if '申報方式' in v:
                    Taxreturnchannel = v.replace('申報方式','').replace('： ','')
                if '申報平台' in v:
                    Taxreturnplat = v.replace('申報平台','').replace('： ','')
                if '首次申報日期' in v:
                    taxreturndate = v.replace('首次申報日期','').replace('： ','')
                if '完成申報時間' in v:
                    taxreturndate = v.replace('完成申報時間','').replace('： ','')
                if '稽徵機關名稱' in v:
                    taxOffice = v.replace('稽徵機關名稱','').replace('： ','')
                if '稽徵機關地址' in v:
                    taxOfficeAddr = v.replace('稽徵機關地址','').replace('： ','')
                if '稽徵機關電話' in v:
                    taxOfficePhone = v.replace('稽徵機關電話','').replace('： ','')
                if '查詢電話' in v:
                    taxOfficePhone = v[v.find('查詢電話:')+len('查詢電話:'):]
                    taxOffice = v[:v.find('查詢電話')]
        else :

            info = soup.find('fieldset').find('p').text
            
            status = 'done'
            INQCode,Taxreturnchannel,Taxreturnplat,taxreturndate,taxOffice,taxOfficeAddr,taxOfficePhone = '','','','','','',''
    
        updatesql(server,username,password,database,status,info,INQCode,Taxreturnchannel,Taxreturnplat,taxreturndate,taxOffice,taxOfficeAddr,taxOfficePhone,entitytype,psid,pid)
        driver.find_element_by_xpath('//*[@id="cmxform"]/p[3]/input[1]').click()
except:
    q = 30
    driver.close()
    import os
    os.system(rf"C:\Py_Project\project\TaxReturn01\TaxReturn_01.bat")
obs = src_obs(server,username,password,database,fromtb,totb,entitytype)
if obs ==0:
    import os
    import sys
    driver.close()
    sys.exit()