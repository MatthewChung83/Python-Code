import io
import time
import datetime
import calendar
import requests
import pyautogui
from PIL import Image
from bs4 import BeautifulSoup
from selenium import webdriver
import selenium.common.exceptions
import pyscreenshot as ImageGrab
import sys
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.expected_conditions import alert_is_present
import urllib.request
import urllib.error
import import_ipynb

from config import *
from etl_func import *

server,database,username,password = db['server'],db['database'],db['username'],db['password']
fromtb,totb = db['fromtb'],db['totb']
url,Drvfile,disk = wbinfo['url'],wbinfo['Drvfile'],wbinfo['disk']

driver = webdriver.Chrome(Drvfile)
driver.get(url)

obs = src_obs(server,username,password,database,fromtb,totb,entitytype)


for i in foo(-1,obs-1):
    try:
        #data =urllib.request.urlopen(url)
        #print(data)
    #except:
        #time.sleep(3)  
        #driver.refresh()
        #time.sleep(3)   
    
        src = dbfrom(server,username,password,database,fromtb,entitytype)
        psid = src[0][0]
        pid = src[0][1].upper().strip()#大寫調整&前後去空值
        cname = src[0][2].strip()#前後去空值
        ci = src[0][3]
        type = src[0][4]
        item,page = '',''
        status = 'N'
        pid_1=pid[1:]
        category = pid_1.isnumeric()
        if len(pid)==8 and type =='C'  :
            driver.find_element(By.CSS_SELECTOR, ".col-xs-12:nth-child(1) input:nth-child(1)").click()
            driver.find_element(By.ID, "queryDebtorNo").clear()
            driver.find_element(By.ID, "queryDebtorNo").send_keys(pid)
        elif len(pid) == 10 and type == 'N' :
            driver.find_element(By.CSS_SELECTOR, ".col-xs-12:nth-child(1) input:nth-child(2)").click()
            try:
                driver.find_element(By.ID, "queryDebtorName").send_keys(cname)
            except:
                #driver.find_element(By.ID, "queryDebtorName").send_keys("")
                #continue
                print(i,psid,pid,cname,'姓名難字')
                updateSQL(server,username,password,database,fromtb,psid,pid,ci,'X',entitytype)
                continue
    
            driver.find_element(By.ID, "queryDebtorName").clear()
            driver.find_element(By.ID, "queryDebtorNo").clear()
            #20211002 Lillian 取消註解+寫Log更新db，使查找繼續E
    
            driver.find_element(By.ID, "queryDebtorName").send_keys(cname)
            driver.find_element(By.ID, "queryDebtorNo").send_keys(pid)
        elif category == False and type == 'P' or len(pid) == 10:
            driver.find_element(By.CSS_SELECTOR, ".col-xs-12:nth-child(1) input:nth-child(3)").click()
            try:
                driver.find_element(By.ID, "queryDebtorName").send_keys(cname)
            except:
                #driver.find_element(By.ID, "queryDebtorName").send_keys("")
                #continue
                print(i,psid,pid,cname,'姓名難字')
                updateSQL(server,username,password,database,fromtb,psid,pid,ci,'X',entitytype)
                continue
    
            driver.find_element(By.ID, "queryDebtorName").clear()
            #20211002 Lillian 取消註解+寫Log更新db，使查找繼續E
    
            driver.find_element(By.ID, "queryDebtorName").send_keys(cname)
            driver.find_element(By.ID, "queryDebtorNo").send_keys(pid)
            
        else :
            import pymssql
            conn = pymssql.connect(server=server, user=username, password=password, database = database)
            cursor = conn.cursor()
        
            script = f"""
            update PropertySecuredtb set type_query = 'err'+'{entitytype}'
            where psid={psid} and pid = '{pid}' and ci = {ci} ;
            """
            cursor.execute(script)
            conn.commit()
            cursor.close()
            conn.close()   
            continue
        driver.find_element(By.XPATH, "(//input[@id=\'button\'])[2]").click()
        
        # 判斷:是否有 result 查詢結果的物件生成
        try:
            element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "result")))
            pageSource = driver.page_source
            soup = BeautifulSoup(pageSource,'html.parser')
            result = soup.find(id="result").text
            
            # 無動產案件
            if '無符合案件' in result:            
                print(i,psid,pid,cname,'無符合案件')
                updateSQL(server,username,password,database,fromtb,psid,pid,ci,status,entitytype)
                driver.find_element(By.ID, "queryDebtorName").clear()
                driver.find_element(By.ID, "queryDebtorNo").clear()
                continue
            
            # 動產案件    
            else:
                pass
        
        finally:
            pass
        # 防呆:Pass輸入資料錯誤跳出的警告視窗
        try:
            WebDriverWait(driver, 3).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            alert.accept()
            driver.find_element(By.ID, "queryDebtorName").clear()
            driver.find_element(By.ID, "queryDebtorNo").clear()
            updateSQL(server,username,password,database,fromtb,psid,pid,ci,status,entitytype)
            continue
        except TimeoutException:
            pass
                
        soup = BeautifulSoup(driver.page_source,'html.parser')
        pageitem = soup.find(class_ = 'row padding-sides25-20').find(class_ = 'col-xs-12 col-sm-6').text.replace('\n','').replace('\t','')
        item = int(pageitem[pageitem.find('共')+1:pageitem.find('筆')])
        page = int(pageitem[pageitem.find('筆')+1:pageitem.find('頁')])
        driver.maximize_window()
        
        printscreen = ''
        for j in range(page):
            
            soup = BeautifulSoup(driver.page_source,'html.parser')
            obs = soup.find('tbody').select('tr')
            
            for k in range(len(obs)):
                obs = soup.find('tbody').find_all('tr')[k]
                item = obs.find_all('td')[0].text ## 項次
                status = obs.find_all('td')[6].text ## 案件狀態
                
                # XPATH 選取案件明細連結 ... 僅一案 & 兩案以上 在 第一案的xpath不同，特別處理
                if len(obs) == 1:
                    click_tag = "//*[@id='"+"table_result']/tbody/tr"
                else:
                    click_tag = "//*[@id='"+"table_result']/tbody/tr["+str(k+1)+"]"
                    
                driver.find_element(By.XPATH,(click_tag)).click()
                
                # 進入案件明細，顯性等待讀取完畢
                try:
                    element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "formInput")))
                    obs_case = BeautifulSoup(driver.page_source,'html.parser')
                    
                    base_scope = obs_case.find_all(id = 'formInput')[0].find(class_ = 'well pubDetailWell center-block')
                    base = base_scope.find_all(class_= 'row pubDetailRow')
                    register_organ = base[0].find_all('div')[1].text ## 登記機關
                    class_ = base[0].find_all('div')[3].text ## 案件類別
                    register_no = base[1].find_all('div')[1].text ## 登記編號
                    register_date = base[1].find_all('div')[3].text ## 登記核准日期
                    change_no = base[2].find_all('div')[1].text ## 變更文號
                    change_date = base[2].find_all('div')[3].text ## 變更核准日期
                    logout_no = base[3].find_all('div')[1].text ## 註銷文號
                    logout_date = base[3].find_all('div')[3].text ## 註銷日期
                    
                    debt_scope = obs_case.find_all(id = 'formInput')[0].find(class_ = 'pubDetailWellTransparent center-block')
                    debt = debt_scope.find_all(class_= 'row pubDetailRow')
                    
                    debt1 = debt[0].find_all('div')[1].text ## 債務人名稱
                    debt2 = debt[1].find_all('div')[1].text ## 債務人統編
                    agent_cname = debt[2].find_all('div')[1].text ## 債務代理人名稱
                    agent_pid = debt[3].find_all('div')[1].text ## 債務代理人統編
                    creditor_name = debt[4].find_all('div')[1].text ## 債權人名稱
                    creditor_id = debt[5].find_all('div')[1].text ## 債權人統編
                    creditor_agent_name = debt[6].find_all('div')[1].text ## 債權代理人名稱
                    creditor_agent_id = debt[7].find_all('div')[1].text ## 債權代理人統編
                    
                    estate_scope = obs_case.find_all(class_ = 'well pubDetailWell center-block')
                    if len(estate_scope) == 2:
                        estate_scope = obs_case.find_all(id = 'formInput')[0].find_all(class_ = 'well pubDetailWell center-block')[1]
                    else:
                        estate_scope = obs_case.find_all(id = 'formInput')[0].find_all(class_ = 'well pubDetailWell center-block')[2]
                    
                    estate = estate_scope.find_all(class_= 'row pubDetailRow')
                    
                    contract_start_date = estate[0].find_all('div')[1].text ## 契約啟始日期
                    contract_end_date = estate[0].find_all('div')[3].text ## 契約終止日期
                    subject_owner_name = estate[1].find_all('div')[1].text ## 標的物所有人名稱
                    estate_amt = estate[1].find_all('div')[3].text.replace('\n','').replace('\t','') ## 擔保債權金額
                    subject_owner_id = estate[2].find_all('div')[1].text ## 標的物所有人統編
                    estate_items = estate[2].find_all('div')[3].text.replace('\n','').replace('\t','') ## 動產明細項數
                    subject_address = estate[3].find_all('div')[1].text.replace('\n','').replace('\t','').replace('\u3000','') ## 標的物所在地
    
                    limitation_flg = estate[4].find_all('div')[1].text.replace('\n','').replace('\t','') ## 是否最高限額
                    float_flg = estate[4].find_all('div')[3].text.replace('\n','').replace('\t','') ## 是否為浮動擔保
                    subject_species = estate[5].find_all('div')[1].text ## 標的物種類
                    
                    datetime_dt = datetime.datetime.today() # 獲得當地時間
                    data_date = datetime_dt.strftime("%Y/%m/%d %H:%M:%S") # 格式化日期
    
                    driver.execute_script("document.body.style.transform='scale(0.75)'")
                    im = ImageGrab.grab(bbox=(0,0,2500,1035))
                    printscreen = rf'{entitytype}_{psid}_{pid}_{ci}_{register_no}.png'
                    #im.save(rf'C:\Py_Project\project\MovableProperty\file\{printscreen}')
                    im.save(rf'C:\Py_Project\project\MovableProperty\file\{entitytype}\{printscreen}')
    
                    docs = (psid,
                            pid,
                            ci,
                            cname,        # '債務人姓名',
                            item,           # '項次',
                            status,           # '案件狀態',
                            register_organ,        # '登記機關',
                            class_,        # '案件類別',
                            register_no,        # '登記編號',
                            register_date,        # '登記核准日期',
                            change_no,        # '變更文號',
                            change_date,        # '變更核准日期',
                            logout_no,        # '註銷文號',
                            logout_date,        # '註銷日期',
                            agent_cname,        # '債務代理人名稱',
                            agent_pid,        # '債務代理人統編',
                            creditor_name,        # '債權人名稱',
                            creditor_id,        # '債權人統編',
                            creditor_agent_name,        # '債權代理人名稱',
                            creditor_agent_id,        # '債權代理人統編',
                            contract_start_date,      # '契約啟始日期',
                            contract_end_date,      # '契約終止日期',
                            subject_owner_name,      # '標的物所有人名稱',
                            estate_amt,      # '擔保債權金額',
                            subject_owner_id,      # '標的物所有人統編',
                            estate_items,      # '動產明細項數',
                            subject_address,      # '標的物所在地',
                            limitation_flg,      # '是否最高限額',
                            float_flg,      # '是否為浮動擔保',
                            subject_species,     # '標的物種類'
                            data_date,  #  資料更新時間
                            printscreen,
                            entitytype,
                           )
                    #driver.execute_script("document.body.style.transform='scale(0.75)'")
                    #im = ImageGrab.grab(bbox=(0,0,2500,1035))
                    #printscreen = rf'{entitytype}_{psid}_{pid}_{ci}_{register_no}.png'
                    #im.save(rf'C:\Py_Project\project\MovableProperty\file\{printscreen}')
                    mp_result = mp_etl(docs)
                    if target_obs(server,username,password,database,totb,psid,pid,ci,register_no,change_no,entitytype) != 0:
                        pass
                    else:
                        toSQL(mp_result, totb, server, database, username, password)
                    # 寫完後退回上一頁
                    driver.back()
                    
                finally:
                    pass
            try:
                driver.find_element(By.XPATH, '//*[@id="queryForm"]/div/div[4]/div[2]/a[3]').click()
            except:
                pass
                #driver.refresh()
                #driver.find_element(By.XPATH, '//*[@id="queryForm"]/div/div[4]/div[2]/a[3]').click()
        status = 'Y'
        updateSQL(server,username,password,database,fromtb,psid,pid,ci,status,entitytype)
        driver.find_element(By.ID, "queryDebtorName").clear()
        driver.find_element(By.ID, "queryDebtorNo").clear()
    except:
        os.system('pkill chrome')
        os.system("kill $(ps aux | grep webdriver| awk '{print $2}')")
        os.system('taskkill /F /IM chrome.exe')
        
        os.system(rf"C:\Py_Project\project\pytest\pybatch_MovableProperty.bat")
    
if src_obs(server,username,password,database,fromtb,totb,entitytype)==0 :
   path = rf'C:\Py_Project\project\MovableProperty\file\Done.txt'
   f = open(path, 'w')
   print('成功執行!!')
   sys.exit()
else :
   pass
        