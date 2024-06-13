from bs4 import BeautifulSoup
import requests
import ddddocr
import time
import re
import sys

from etl_func import *
from config import *
from dict import *

# 宣告資料庫、網站、迭代...等參數
server,database,username,password,totb1 = db['server'],db['database'],db['username'],db['password'],db['totb1']
url,captchaImg = wbinfo['url'],wbinfo['captchaImg']
imgp = pics['imgp']
q = 0

# 取名單數量
obs = src_obs(server,username,password,database,totb1)
src = dbfrom(server,username,password,database,totb1)
print(obs)
for i in range(len(src)):

    # 逐筆取庫內資料
    

    ID = re.sub(r"\s+", "", src[i][3])
    birthday = ('0'+re.sub(r"\s+", "", src[i][4]).replace('/',''))[-7:]

    resp = capcha_resp(url,captchaImg,imgp,q,ID,birthday)

    # 判斷查找結果之分類依據
    driver_info = resp.find(class_ = 'tb_list_std')
    # 查回相關資訊給定預設值
    driver_type,driver_status,DRvaliddate,status,updatetime = '','','','N',(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    if driver_info:
        if len(driver_info) == 5:
            driver_status = resp.find('tbody').find('td').text
        else:
            driver_type = re.sub(r"\s+", "", driver_info.find_all('tr')[1].find_all('td')[0].text)
            driver_status = re.sub(r"\s+", "", driver_info.find_all('tr')[1].find_all('td')[1].text)
            DRvaliddate = re.sub(r"\s+", "", driver_info.find_all('tr')[1].find_all('td')[2].text)
            if '死亡' in driver_status:
                status = 'Y'
    else:
        driver_status = '查無汽機車駕照'
    
    # 更新查回結果至資料庫
    updatetime = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    updateSQL(server,username,password,database,totb1,status,updatetime,ID,driver_type,driver_status,DRvaliddate)
    print(ID,driver_type,driver_status,DRvaliddate,status)
    ex_obs = exit_obs(server,username,password,database,totb1)
    if ex_obs >= 50000:
        sys.exit()
    else:
        pass
    
    