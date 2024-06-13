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
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import ddddocr
import io

imgp = r'C:\Py_Project\project\tax_refund\captcha_04.jpg'
# 讀取的sql來源
def fromsql(host,user,password,database,src_tb):
    conn = pymssql.connect(host = host,user = user,password = password,database = database)
    cursor = conn.cursor(as_dict=True)
    script = f"""select * from {src_tb} where info is null and type = 'ONHAND-20230821_04' order by pid """
    cursor.execute(script)
    sql_src = []
    for row in cursor:
        sql_src.append(row)
    cursor.close()
    conn.close()
    return sql_src

# 更新的sql 目的資料表
def updatesql(host,user,password,database,tar_tb,info,psid,pid):
    conn = pymssql.connect(host = host,user = user,password = password,database = database)
    cursor = conn.cursor(as_dict=True)
    script = """update {} set info = '{}' where psid = {} and pid = '{}'""".format(tar_tb,info,psid,pid)
    print(script)
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    
# 確認元素是否存在，存在返回flag=true，否則返回false
def isElementExist(element):
    flag=True
    try:
        driver.find_element_by_xpath(element)
        return flag
    except:
        flag=False
        return flag
    
# 引入辨識模型
#clf = models.load_model(r'.\model\etax_nat_query_model.h5')

# 引入ChromeDriver
DriverPath = rf'C:\Py_Project\env\chromedriver_win32\chromedriver'
#中區Url = 'https://www.etax.nat.gov.tw/etwmain/etw133w1/b01'
#北區Url = 'https://www.etax.nat.gov.tw/etwmain/etw133w1/a05'
#南區Url = 'https://www.etax.nat.gov.tw/etwmain/etw133w1/d01'
Url = 'https://www.etax.nat.gov.tw/etwmain/etw133w1/e01'
chrome_options = Options()
chrome_options.add_argument('--headless')
chromedriver = DriverPath
driver = webdriver.Chrome(chromedriver,chrome_options=chrome_options)
driver.get(Url)
#driver.minimize_window()

src = fromsql('10.10.0.94','CLUSER','Ucredit7607','CL_Daily','taxrefundtb')
obs = len(src)




def foo(num):
    while num < obs:
        num = num + 1 
        yield num

#for i in range(len(src)):
for i in range(obs):
    try:
        x=1
        while x < 10:
            psid = src[i]['psid']
            pid = src[i]['pid']
            result = ''
            driver.find_element_by_xpath('//*[@id="userIdnBan"]').send_keys(pid)
            #驗證碼
            time.sleep(0.5)
            img_ele = driver.find_element_by_xpath('//*[@id="queryForm"]/div[2]/div[1]/div/div/div/etw-captcha/div/img')
            img = Image.open( io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
            img.save(imgp)
            ocr = ddddocr.DdddOcr()
            with open(imgp, 'rb') as f:
                img_bytes = f.read()
            res = ocr.classification(img_bytes)
            time.sleep(0.5)
            driver.find_element_by_xpath('//*[@id="captchaText"]').send_keys(res)
            time.sleep(0.5)
            driver.find_element_by_xpath('//*[@id="queryForm"]/div[2]/div[2]/div/button').click()
            try:
                #driver.find_element_by_xpath('/html/body/ngb-modal-window/div/div/jhi-dialog/div/div[2]').text
                time.sleep(3)
                driver.find_element_by_xpath('/html/body/ngb-modal-window/div/div/jhi-dialog/div/div[3]/div/button').click()
                driver.refresh()
                x = x+1
            except:
                x = 10
                pass
        time.sleep(3)
        info = driver.find_element_by_xpath('//*[@id="resultArea"]/div/div/table/tbody/tr/td').text.replace('\n','')
        insertdate = dt.today().strftime("%Y/%m/%d %H:%M:%S")
        updatesql('10.10.0.94','CLUSER','Ucredit7607','CL_Daily','taxrefundtb',info,psid,pid)
        print(i,info,psid,pid,insertdate)
        driver.refresh()
    except:
        driver.refresh()
