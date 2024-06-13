import pytest
import time
import json
import pandas as pd
import csv
import pymssql
import datetime

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import TimeoutException

def ReadCsv(file_path,datalist):
    csvFileToWrite = open(file_path, 'a', encoding='utf-8-sig', newline='')
    csvDataToWrite = csv.writer(csvFileToWrite)
    csvDataToWrite.writerow(datalist)
    csvFileToWrite.close()

url = 'https://ppstrq.nat.gov.tw/pps/pubQuery/PropertyQuery/propertyQuery.do'
DriverPath = r'C:\Py_Project\env\chromedriver_win32\chromedriver'

offobs = 300000
fetobs = 10000
start = 9940

title = (
         'psid',
         'pid',
         'ci',
         'cname', #'債務人姓名',
         'item', #'項次',
         'status', #'案件狀態',
         'register_organ', #'登記機關',
         'class', #'案件類別',
         'register_no', #'登記編號',
         'register_date', #'登記核准日期',
         'change_no', #'變更文號',
         'change_date', #'變更核准日期',
         'logout_no', #'註銷文號',
         'logout_date', #'註銷日期',
         'agent_cname', #'債務代理人名稱',
         'agent_pid', #'債務代理人統編',
         'creditor_name', #'債權人名稱',
         'creditor_id', #'債權人統編',
         'creditor_agent_name', #'債權代理人名稱',
         'creditor_agent_id', #'債權代理人統編',
         'contract_start_date', #'契約啟始日期',
         'contract_end_date', #'契約終止日期',
         'subject_owner_name', #'標的物所有人名稱',
         'estate_amt', #'擔保債權金額',
         'subject_owner_id', #'標的物所有人統編',
         'estate_items', #'動產明細項數',
         'subject_address', #'標的物所在地',
         'limitation_flg', #'是否最高限額',
         'float_flg', #'是否為浮動擔保',
         'subject_species',  #'標的物種類'
         'data_date' # 資料更新時間
        )


CsvToPath = rf'C:\Py_Project\output\Estate_output_{offobs}to{fetobs}.csv'
ReadCsv(CsvToPath,title)

driver = webdriver.Chrome(DriverPath)
driver.get(url)

conn = pymssql.connect(server='VNT07', user='.\chris', password='Ucs28289788')
cursor = conn.cursor()
cursor.execute(f"SELECT * FROM [UIS].[dbo].[PropertySecuredtb] ORDER BY psid OFFSET {offobs} ROWS FETCH NEXT {fetobs} ROWS ONLY")
src = cursor.fetchall()
cursor.close()
conn.close()

for i in range(start,len(src)):
    psid = src[i][0]
    pid = src[i][1]
    cname = src[i][2]
    ci = src[i][4]

    if '' in cname or '' in cname or '' in cname:
        continue
    
    if len(pid) == 8:
        driver.find_element(By.CSS_SELECTOR, ".col-xs-12:nth-child(1) input:nth-child(1)").click()
        driver.find_element(By.ID, "queryDebtorNo").send_keys(pid)
    else:
        driver.find_element(By.CSS_SELECTOR, ".col-xs-12:nth-child(1) input:nth-child(2)").click()
        driver.find_element(By.ID, "queryDebtorName").send_keys(cname)
        driver.find_element(By.ID, "queryDebtorNo").send_keys(pid)
        
    driver.find_element(By.XPATH, "(//input[@id=\'button\'])[2]").click()
    
    # 防呆:Pass輸入資料錯誤跳出的警告視窗
    try:
        WebDriverWait(driver, 3).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
        driver.find_element(By.ID, "queryDebtorName").clear()
        driver.find_element(By.ID, "queryDebtorNo").clear()
        continue
    except TimeoutException:
        pass
    
    # 判斷:是否有 result 查詢結果的物件生成
    try:
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "result")))
        pageSource = driver.page_source
        soup = BeautifulSoup(pageSource,'html.parser')
        result = soup.find(id="result").text
        
        # 無動產案件
        if '無符合案件' in result:
            
            print(i,psid,pid,cname,'無符合案件')
        
        # 動產案件    
        else:
            
            soup = BeautifulSoup(driver.page_source,'html.parser')
            obs = soup.find('tbody').select('tr')
            head = soup.find('thead').select('th')
            
            # 每個ID項下案件數
            for j in range(len(obs)):
                
                # 取案件摘要
                obs = soup.find('tbody').find_all('tr')[j]
                h1 = obs.find_all('td')[0].text ## 項次
                h7 = obs.find_all('td')[6].text ## 案件狀態
                
                # XPATH 選取案件明細連結 ... 僅一案 & 兩案以上 在 第一案的xpath不同，特別處理
                if len(obs) == 1:
                    click_tag = "//*[@id='"+"table_result']/tbody/tr"
                else:
                    click_tag = "//*[@id='"+"table_result']/tbody/tr["+str(j+1)+"]"
                
                driver.find_element_by_xpath(click_tag).click()
                
                # 進入案件明細，顯性等待讀取完畢
                try:
                    element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "formInput")))
                    obs_case = BeautifulSoup(driver.page_source,'html.parser')
                    
                    base_scope = obs_case.find_all(id = 'formInput')[0].find(class_ = 'well pubDetailWell center-block')
                    base = base_scope.find_all(class_= 'row pubDetailRow')
                
                    base1 = base[0].find_all('div')[1].text ## 登記機關
                    base2 = base[0].find_all('div')[3].text ## 案件類別
                    base3 = base[1].find_all('div')[1].text ## 登記編號
                    base4 = base[1].find_all('div')[3].text ## 登記核准日期
                    base5 = base[2].find_all('div')[1].text ## 變更文號
                    base6 = base[2].find_all('div')[3].text ## 變更核准日期
                    base7 = base[3].find_all('div')[1].text ## 註銷文號
                    base8 = base[3].find_all('div')[3].text ## 註銷日期
                    
                    debt_scope = obs_case.find_all(id = 'formInput')[0].find(class_ = 'pubDetailWellTransparent center-block')
                    debt = debt_scope.find_all(class_= 'row pubDetailRow')
                    
                    debt1 = debt[0].find_all('div')[1].text ## 債務人名稱
                    debt2 = debt[1].find_all('div')[1].text ## 債務人統編
                    debt3 = debt[2].find_all('div')[1].text ## 債務代理人名稱
                    debt4 = debt[3].find_all('div')[1].text ## 債務代理人統編
                    debt5 = debt[4].find_all('div')[1].text ## 債權人名稱
                    debt6 = debt[5].find_all('div')[1].text ## 債權人統編
                    debt7 = debt[6].find_all('div')[1].text ## 債權代理人名稱
                    debt8 = debt[7].find_all('div')[1].text ## 債權代理人統編
                    
                    estate_scope = obs_case.find_all(class_ = 'well pubDetailWell center-block')
                    if len(estate_scope) == 2:
                        estate_scope = obs_case.find_all(id = 'formInput')[0].find_all(class_ = 'well pubDetailWell center-block')[1]
                    else:
                        estate_scope = obs_case.find_all(id = 'formInput')[0].find_all(class_ = 'well pubDetailWell center-block')[2]
                    
                    estate = estate_scope.find_all(class_= 'row pubDetailRow')
                    
                    estate1 = estate[0].find_all('div')[1].text ## 契約啟始日期
                    estate2 = estate[0].find_all('div')[3].text ## 契約終止日期
                    estate3 = estate[1].find_all('div')[1].text ## 標的物所有人名稱
                    estate4 = estate[1].find_all('div')[3].text.replace('\n','').replace('\t','') ## 擔保債權金額
                    estate5 = estate[2].find_all('div')[1].text ## 標的物所有人統編
                    estate6 = estate[2].find_all('div')[3].text.replace('\n','').replace('\t','') ## 動產明細項數
                    estate7 = estate[3].find_all('div')[1].text.replace('\n','').replace('\t','').replace('\u3000','') ## 標的物所在地

                    estate8 = estate[4].find_all('div')[1].text.replace('\n','').replace('\t','') ## 是否最高限額
                    estate9 = estate[4].find_all('div')[3].text.replace('\n','').replace('\t','') ## 是否為浮動擔保
                    estate10 = estate[5].find_all('div')[1].text ## 標的物種類
                    
                    datetime_dt = datetime.datetime.today() # 獲得當地時間
                    datetime_str = datetime_dt.strftime("%Y/%m/%d %H:%M:%S") # 格式化日期
                    
                    data = (psid,
                            pid,
                            ci,
                            cname,        # '債務人姓名',
                            h1,           # '項次',
                            h7,           # '案件狀態',
                            base1,        # '登記機關',
                            base2,        # '案件類別',
                            base3,        # '登記編號',
                            base4,        # '登記核准日期',
                            base5,        # '變更文號',
                            base6,        # '變更核准日期',
                            base7,        # '註銷文號',
                            base8,        # '註銷日期',
                            debt3,        # '債務代理人名稱',
                            debt4,        # '債務代理人統編',
                            debt5,        # '債權人名稱',
                            debt6,        # '債權人統編',
                            debt7,        # '債權代理人名稱',
                            debt8,        # '債權代理人統編',
                            estate1,      # '契約啟始日期',
                            estate2,      # '契約終止日期',
                            estate3,      # '標的物所有人名稱',
                            estate4,      # '擔保債權金額',
                            estate5,      # '標的物所有人統編',
                            estate6,      # '動產明細項數',
                            estate7,      # '標的物所在地',
                            estate8,      # '是否最高限額',
                            estate9,      # '是否為浮動擔保',
                            estate10,     # '標的物種類'
                            datetime_str  #  資料更新時間
                           )
                    
                    # 將案件明細，寫入csv檔案
                    CsvToPath = rf'C:\Py_Project\output\Estate_output_{offobs}to{fetobs}.csv'
                    ReadCsv(CsvToPath,data)
                    print(i,data)

                    # 寫完後退回上一頁
                    driver.back()
                finally:
                    pass
    finally:
        pass

    # 清除先前填入ID等資訊，continue for loop
    driver.find_element(By.ID, "queryDebtorName").clear()
    driver.find_element(By.ID, "queryDebtorNo").clear()