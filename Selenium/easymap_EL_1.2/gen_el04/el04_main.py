import io
import time
import calendar
import requests
import os
from PIL import Image
from bs4 import BeautifulSoup
from selenium import webdriver
import selenium.common.exceptions

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.expected_conditions import alert_is_present



from config import *
from etl_func import *
from batch_args import *
from OCR_Mod import *

import os

os.system('pkill chrome')
os.system("kill $(ps aux | grep webdriver| awk '{print $2}')")
os.system('taskkill /F /IM chrome.exe')


Machine,server,database,username,password = hostname(),db['server'],db['database'],db['username'],db['password']
servicetb,fromtb,totb,entitytype = db['servicetb'],db['fromtb'],db['totb'],db['entitytype']
MGtb,fromtb,totb,todtltb,EntityPhase,EntityPhase_next,EntityPath,EntityPath_next = db['MGtb'],db['fromtb'],db['totb'],db['todtltb'],db['EntityPhase'],db['EntityPhase_next'],db['EntityPath'],db['EntityPath_next']
acc_name,acc_pwd,url,Drvfile = wbinfo['acc_name'],wbinfo['acc_pwd'],wbinfo['url'],wbinfo['Drvfile']
Apurl,imgf,imgp = APinfo['Apurl'],APinfo['imgf'],APinfo['imgp']

driver = webdriver.Chrome(Drvfile)
driver.get(url)

#createtmp(server,username,password,database,fromtb,totb,entitytype)
obs = src_obs(server,username,password,database,fromtb,totb,entitytype)
entrylog = ''

### [進入網站] 勾選同意 ###
try:
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID,'ok')))
    Weblog = '(ok--Access)'
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID,'ok')))
    driver.switch_to.default_content()
    Weblog = '(ok switch frame--Access)'
except:
    Weblog = '(ok--Fail)'
entrylog = entrylog + Weblog

if '(ok--Access)' in entrylog:
    driver.find_element_by_id('ok').click()
elif '(ok switch frame--Access)' in entrylog:
    driver.switch_to.frame('untie')
    driver.find_element(By.ID,'ok').click()
    driver.switch_to.default_content()
else:
    driver.quit()

### [進入網站] 同意進入 ###
try:
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID,'yes')))
    Weblog = '(yes--Access)'
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID,'yes')))
    driver.switch_to.default_content()
    Weblog = '(yes switch frame--Access)'
except:
    Weblog = '(yes--Fail)'
entrylog = entrylog + Weblog

if '(yes--Access)' in entrylog:
    driver.find_element_by_id('yes').click()
elif '(yes switch frame--Access)' in entrylog:
    driver.switch_to.frame('untie')
    driver.find_element_by_id('yes').click()
    driver.switch_to.default_content()
else:
    driver.quit()

### [進入網站] 登入帳號 ###
try:
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID,'Img2')))
    Weblog = '(Img2--Access)'
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID,'Img2')))
    driver.switch_to.default_content()
    Weblog = '(Img2 switch frame--Access)'
except:
    Weblog = '(Img2--Fail)'
entrylog = entrylog + Weblog

if '(Img2--Access)' in entrylog:
    driver.find_element_by_id('Img2').click()
elif '(Img2 switch frame--Access)' in entrylog:
    driver.switch_to.frame('untie')
    driver.find_element_by_id('Img2').click()
    driver.switch_to.default_content()
else:
    driver.quit()

q = 0   # 索引變數
while q < 100:    
    ### [進入網站] 輸入帳號 ###
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID,'aa-uid')))
        Weblog = '(aa-uid--Access)'
    except selenium.common.exceptions.TimeoutException:
        driver.switch_to.frame('untie')
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID,'aa-uid')))
        driver.switch_to.default_content()
        Weblog = '(aa-uid switch frame--Access)'
    except:
        Weblog = '(aa-uid--Fail)'
        
    entrylog = entrylog + Weblog
    
    if '(aa-uid--Access)' in entrylog:
        driver.find_element(By.ID, 'aa-uid').send_keys(acc_name)
        
    elif '(aa-uid switch frame--Access)' in entrylog:
        driver.switch_to.frame('untie')
        driver.find_element(By.ID, 'aa-uid').send_keys(acc_name)
        driver.switch_to.default_content()
    
    else:
        driver.quit()
    
    ### [進入網站] 輸入密碼 ###
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID,'aa-passwd')))
        Weblog = '(aa-passwd--Access)'
    except selenium.common.exceptions.TimeoutException:
        driver.switch_to.frame('untie')
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID,'aa-passwd')))
        driver.switch_to.default_content()
        Weblog = '(aa-passwd switch frame--Access)'
    except:
        Weblog = '(aa-passwd--Fail)'
        
    entrylog = entrylog + Weblog
    
    if '(aa-passwd--Access)' in entrylog:
        driver.find_element(By.ID, 'aa-passwd').send_keys(acc_pwd)
        
    elif '(aa-passwd switch frame--Access)' in entrylog:
        driver.switch_to.frame('untie')
        driver.find_element(By.ID, 'aa-passwd').send_keys(acc_pwd)
        driver.switch_to.default_content()
        
    else:
        driver.quit()
    
    ### [進入網站] 解析captcha ###
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID,'AAAIden1')))
        Weblog = '(AAAIden1--Access)'
    except selenium.common.exceptions.TimeoutException:
        driver.switch_to.frame('untie')
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID,'AAAIden1')))
        driver.switch_to.default_content()
        Weblog = '(AAAIden1 switch frame--Access)'
    except:
        Weblog = '(AAAIden1--Fail)'
        
    entrylog = entrylog + Weblog
    
    if '(AAAIden1--Access)' in entrylog:
        img_ele = driver.find_element_by_id('AAAIden1')
        img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
        img.save(imgp)
        
    elif '(AAAIden1 switch frame--Access)' in entrylog:
        driver.switch_to.frame('untie')
        img_ele = driver.find_element_by_id('AAAIden1')
        img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
        img.save(imgp)
        driver.switch_to.default_content()
        
    else:
        driver.quit()
    
    ### [進入網站] 解析及輸入captcha ###
    captchaocr = CAPTCHAOCR(Apurl,imgf,imgp).replace('"','')
    l = len(captchaocr)
    
    while l < 4:
        driver.find_element_by_xpath('//*[@id="AuthScreen"]/div[1]/p[4]/a').click()
        img_ele = driver.find_element_by_id('AAAIden1')
        img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
        img.save(imgp)
        captchaocr = CAPTCHAOCR(Apurl,imgf,imgp).replace('"','')
        l = len(captchaocr)

    ### [進入網站] 輸入captcha ###
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME,'aa-captchaID')))
        Weblog = '(aa-captchaID--Access)'
    except selenium.common.exceptions.TimeoutException:
        driver.switch_to.frame('untie')
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME,'aa-captchaID')))
        driver.switch_to.default_content()
        Weblog = '(aa-captchaID switch frame--Access)'
    except:
        Weblog = '(aa-captchaID--Fail)'
    
    entrylog = entrylog + Weblog
    
    if '(aa-captchaID--Access)' in entrylog:
        driver.find_element_by_name('aa-captchaID').send_keys(captchaocr)
        time.sleep(5)
    
    elif '(aa-captchaID switch frame--Access)' in entrylog:
        driver.switch_to.frame('untie')
        driver.find_element_by_name('aa-captchaID').send_keys(captchaocr)
        time.sleep(5)
        driver.switch_to.default_content()
        
    else:
        driver.quit()        
    
    ### [進入網站] 進入網站 ###
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID,'submit_hn')))
        Weblog = '(submit_hn--Access)'
    except selenium.common.exceptions.TimeoutException:
        driver.switch_to.frame('untie')
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID,'submit_hn')))
        driver.switch_to.default_content()
        Weblog = '(submit_hn switch frame--Access)'
    except:
        Weblog = '(submit_hn--Fail)'
    
    entrylog = entrylog + Weblog
    
    if '(submit_hn--Access)' in entrylog:
        driver.find_element(By.ID, 'submit_hn').click()
    
    elif '(submit_hn switch frame--Access)' in entrylog:
        driver.switch_to.frame('untie')
        driver.find_element(By.ID, 'submit_hn').click()
        driver.switch_to.default_content()
        
    else:
        driver.quit()        
    
    try:
        WebDriverWait(driver, 3).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        time.sleep(5)
        alert.accept()
        time.sleep(2)
        print("alert accepted")
    
    except TimeoutException:
        print("no alert")
        break
        
    q += 1

for i in foo(-1,obs-1):
    
    time.sleep(1)
    el_info = dbfrom(server,username,password,database,fromtb,totb,entitytype)
    eaid,pid,addressi,city,area = el_info[0][0],el_info[0][1],el_info[0][2],el_info[0][4].split(' ',1)[0],el_info[0][4].split(' ',1)[1]
    print(eaid)
    lot_code,landno = el_info[0][5][:4],el_info[0][6]
    appstatus,el_number,el_appdttm,app_result = '','','',''
    # 檢核發查時間
    apply_year = int(time.strftime("%Y", time.localtime()))
    apply_month = int(time.strftime("%m", time.localtime()))
    apply_day = int(time.strftime("%d", time.localtime()))
    apply_time = int(time.strftime("%H%M%S"))
    apply_dttm = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    svcweekday = calendar.weekday(apply_year,apply_month,apply_day) + 1
    svctime = servicetime(server,username,password,database,servicetb,city,svcweekday)
    input_log = ''
    
    
    appsuccessobs = APPSUCCESSOBS(server,username,password,database,fromtb,totb,entitytype)
    insertMG(server,username,password,database,MGtb,entitytype,Machine,EntityPhase_next,totb,appsuccessobs,'null',apply_dttm,EntityPath_next)
    
    if svctime == []:
        timelog = '(out service period)'
        
    elif svctime[0][2] < apply_time < svctime[0][3]:
        timelog = '(' + str(apply_time) + ' in service period)'
        
    else:
        timelog = '(' + str(apply_time) + ' out service period)'
        
    input_log = input_log + timelog

    # 檢核地址合規
    if city != city or area != area:
        data_log = '(address error)'
    else:
        data_log = '(address access)'
    
    input_log = input_log + data_log
    
    if 'error' in input_log:
        appstatus,app_result = 'Y','N'
        appdocs = (eaid,pid,addressi,landno,purpose,appstatus,entitytype,el_number,el_appdttm,app_result,insertdate,input_log)
        EL04_tmp_result = EL04_tmp_etl(appdocs)
        toSQL(EL04_tmp_result, totb, server, database, username, password)
        appobs = APPOBS(server,username,password,database,totb,entitytype)
        
        if i == (obs - 1) and appobs < obs:
            status = 'Re-Exec'
        
        elif i == (obs - 1):
            status = 'Done'
        
        else: 
            status = 'Continue'

        print(updateMGM(server,username,password,database,MGtb,appobs,status,apply_dttm,entitytype,Machine,EntityPhase,todtltb))
        #driver.refresh()
        continue
        
    if 'out' in input_log:
        appdocs = (eaid,pid,addressi,landno,purpose,appstatus,entitytype,el_number,el_appdttm,app_result,insertdate,input_log)
        EL04_tmp_result = EL04_tmp_etl(appdocs)
        #driver.refresh()
        continue        
    
    # 申請流程

    app_log = ''
    if len(landno) == 8:
        
        # 進入申請頁面
        try:
            driver.switch_to.frame('untie')
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME,'Image14')))
            driver.switch_to.default_content()
            Weblog = '(Image14 switch frame--Access)'

        except selenium.common.exceptions.TimeoutException:
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME,'Image14')))
            Weblog = '(Image14--Access)'

        except:
            Weblog = '(Image14--Fail)'
            
        app_log = app_log + Weblog
        
        if '(Image14--Access)' in app_log:
            driver.find_element(By.NAME, 'Image14').click()
        elif '(Image14 switch frame--Access)' in app_log:
            driver.switch_to.frame('untie')
            driver.find_element(By.NAME, 'Image14').click()
            driver.switch_to.default_content()
        else:
            #driver.refresh()
            continue
        
        soup = BeautifulSoup(driver.page_source,'html.parser')
        entry_output = soup.find_all('h3')
        
        # 選擇城市        
        try:
            driver.switch_to.frame('untie')
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'City_ID')))
            driver.switch_to.default_content()
            Weblog = '(City_ID switch frame--Access)'
            
        except selenium.common.exceptions.TimeoutException:
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'City_ID')))
            Weblog = '(City_ID--Access)'
            
        except:
            Weblog = '(City_ID--Fail)'
            
        app_log = app_log + Weblog
        
        if '(City_ID--Access)' in app_log:
            pass
        elif '(City_ID switch frame--Access)' in app_log:
            driver.switch_to.frame('untie')
        else:
            #driver.refresh()
            continue
        
        soup = BeautifulSoup(driver.page_source,'html.parser')
        
        for c in soup.find_all(id = 'City_ID')[0].find_all('option'):            
            if c.text == city:
                ele_city = Select(driver.find_element_by_id('City_ID'))
                ele_city.select_by_visible_text(city)
                Weblog = '(City_ID input--Access)'
                break
                
            else:
                Weblog = '(City_ID input--Fail)'
        
        app_log = app_log + Weblog
        
        if '(City_ID switch frame--Access)' in app_log:
            driver.switch_to.default_content()
            
        # 選擇鄉鎮市區
        try:
            driver.switch_to.frame('untie')
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'area_id')))
            driver.switch_to.default_content()
            Weblog = '(area_id switch frame--Access)'
            
        except selenium.common.exceptions.TimeoutException:
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'area_id')))
            Weblog = '(area_id--Access)'
            
        except:
            Weblog = '(area_id--Fail)'
            
        app_log = app_log + Weblog
        
        if '(area_id--Access)' in app_log:
            pass
        
        elif '(area_id switch frame--Access)' in app_log:
            driver.switch_to.frame('untie')
        
        elif '(area_id--Fail)' in app_log:
            print('(area_id--Fail)')
            appstatus,app_result = 'Y','N'
            appdocs = (eaid,pid,addressi,landno,purpose,appstatus,entitytype,el_number,el_appdttm,app_result,insertdate,input_log)
            EL04_tmp_result = EL04_tmp_etl(appdocs)
            toSQL(EL04_tmp_result, totb, server, database, username, password)
            driver.find_element(By.NAME, 'Image3').click()
            continue
        
        else:
            #driver.refresh()
            continue
        
        soup = BeautifulSoup(driver.page_source,'html.parser')
        
        for a in soup.find_all(id = 'area_id')[0].find_all('option'):
            if a.text == area:
                ele_area = Select(driver.find_element_by_id('area_id'))
                ele_area.select_by_visible_text(area)
                Weblog = '(area_id input--Access)'
                break
            else:
                Weblog = '(area_id input--Fail)'
                
        app_log = app_log + Weblog
        
        # 輸入申請人統編
        try:
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'CONSIGNOR_ID')))
            Weblog = '(CONSIGNOR_ID--Access)'
            
        except selenium.common.exceptions.TimeoutException:
            driver.switch_to.frame('untie')
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'CONSIGNOR_ID')))
            driver.switch_to.default_content()
            Weblog = '(CONSIGNOR_ID switch frame--Access)'
            
        except:
            Weblog = '(CONSIGNOR_ID--Fail)'
            
        app_log = app_log + Weblog
    
        if '(CONSIGNOR_ID--Access)' in app_log:
            driver.find_element_by_id('CONSIGNOR_ID').send_keys(pid)
    
        elif '(CONSIGNOR_ID switch frame--Access)' in app_log:
            driver.switch_to.frame('untie')
            driver.find_element_by_id('CONSIGNOR_ID').send_keys(pid)
            driver.switch_to.default_content()
            
        else:
            #driver.refresh()
            continue
        
        # 輸入查詢之統編
        try:
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'INPUT_015')))
            Weblog = '(INPUT_015--Access)'
            
        except selenium.common.exceptions.TimeoutException:
            driver.switch_to.frame('untie')
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'INPUT_015')))
            driver.switch_to.default_content()
            Weblog = '(INPUT_015 switch frame--Access)'    
            
        except:
            Weblog = '(INPUT_015--Fail)'
            
        app_log = app_log + Weblog
    
        if '(INPUT_015--Access)' in app_log:
            driver.find_element_by_id('INPUT_015').send_keys(pid)
    
        elif '(INPUT_015 switch frame--Access)' in app_log:
            driver.switch_to.frame('untie')
            driver.find_element_by_id('INPUT_015').send_keys(pid)
            driver.switch_to.default_content()
            
        else:            
            #driver.refresh()
            continue
            
        #輸入申請用途
        try:
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'applyfor')))
            Weblog = '(applyfor--Access)'
            
        except selenium.common.exceptions.TimeoutException:
            driver.switch_to.frame('untie')
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'applyfor')))
            driver.switch_to.default_content()
            Weblog = '(applyfor switch frame--Access)'    
            
        except:
            Weblog = '(applyfor--Fail)'
        app_log = app_log + Weblog
        
        if '(applyfor--Access)' in app_log:
            Select(driver.find_element_by_id('applyfor')).select_by_visible_text('自行參考')
    
        elif '(applyfor switch frame--Access)' in app_log:
            driver.switch_to.frame('untie')
            Select(driver.find_element_by_id('applyfor')).select_by_visible_text('自行參考')
            driver.switch_to.default_content()
            
        else:
            #driver.refresh()
            continue
        
        # 輸入查詢之地段編號
        try:
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'INPUT_011')))
            Weblog = '(INPUT_011--Access)'
            
        except selenium.common.exceptions.TimeoutException:
            driver.switch_to.frame('untie')
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'INPUT_011')))
            driver.switch_to.default_content()
            Weblog = '(INPUT_011 switch frame--Access)'    
            
        except:
            Weblog = '(INPUT_011--Access)'
            
        app_log = app_log + Weblog
    
        if '(INPUT_011--Access)' in app_log:
            driver.find_element_by_id('INPUT_011').clear()
            driver.find_element_by_id('INPUT_011').send_keys(lot_code)
    
        elif '(INPUT_011 switch frame--Access)' in app_log:
            driver.switch_to.frame('untie')
            driver.find_element_by_id('INPUT_011').clear()
            driver.find_element_by_id('INPUT_011').send_keys(lot_code)
            driver.switch_to.default_content()
            
        else:
            #driver.refresh()
            continue
        
        # 輸入查詢之地號
        try:
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'INPUT_013')))
            Weblog = '(INPUT_013--Access)'
            
        except selenium.common.exceptions.TimeoutException:
            driver.switch_to.frame('untie')
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'INPUT_013')))
            driver.switch_to.default_content()
            Weblog = '(INPUT_013 switch frame--Access)'
        
        except:
            Weblog = '(INPUT_013--Fail)'
            
        app_log = app_log + Weblog
            
        if '(INPUT_013--Access)' in app_log:
            driver.find_element_by_id('INPUT_013').send_keys(landno)
    
        elif '(INPUT_013 switch frame--Access)' in app_log:
            driver.switch_to.frame('untie')
            driver.find_element_by_id('INPUT_013').send_keys(landno)
            driver.switch_to.default_content()
            
        else:
            #driver.refresh()
            continue
        
        # 勾選登記謄本
        try:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "INPUT_021")))
            Weblog = '(INPUT_021--Access)'
            
        except selenium.common.exceptions.TimeoutException:
            driver.switch_to.frame('untie')
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "INPUT_021")))
            driver.switch_to.default_content()
            Weblog = '(INPUT_021 switch frame--Access)'
        
        except:
            Weblog = '(INPUT_021--Fail)'
            
        app_log = app_log + Weblog
        
        if '(INPUT_021--Access)' in app_log:
            driver.find_element_by_name('INPUT_021').click()
            purpose = driver.find_element_by_class_name('applyitem').text
    
        elif '(INPUT_021 switch frame--Access)' in app_log:
            driver.switch_to.frame('untie')
            driver.find_element_by_name('INPUT_021').click()
            purpose = driver.find_element_by_class_name('applyitem').text
            driver.switch_to.default_content()
            
        else:
            #driver.refresh()
            continue
            
        # 取消部分內容列印
        try:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "INPUT_022")))
            Weblog = '(INPUT_022--Access)'
            
        except selenium.common.exceptions.TimeoutException:
            driver.switch_to.frame('untie')
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "INPUT_022")))
            driver.switch_to.default_content()
            Weblog = '(INPUT_022 switch frame--Access)'
        
        except:
            Weblog = '(INPUT_022--Fail)'
            
        app_log = app_log + Weblog
        
        if '(INPUT_022--Access)' in app_log:
            driver.find_element_by_name('INPUT_022').click()
    
        elif '(INPUT_022 switch frame--Access)' in app_log:
            driver.switch_to.frame('untie')
            driver.find_element_by_name('INPUT_022').click()
            driver.switch_to.default_content()
        
        else:
            #driver.refresh()
            continue
            
        # 頁面新增資料
        try:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "btnnew")))
            Weblog = '(btnnew--Access)'
            
        except selenium.common.exceptions.TimeoutException:
            driver.switch_to.frame('untie')
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "btnnew")))
            driver.switch_to.default_content()
            Weblog = '(btnnew switch frame--Access)'
        
        except:
            Weblog = '(btnnew--Fail)'
            
        app_log = app_log + Weblog
        
        if '(btnnew--Access)' in app_log:
            driver.find_element_by_name('btnnew').click()
    
        elif '(btnnew switch frame--Access)' in app_log:
            driver.switch_to.frame('untie')
            driver.find_element_by_name('btnnew').click()
            driver.switch_to.default_content()
            
        else:
            #driver.refresh()
            continue
        
        # 送出申請
        try:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "btnsend")))
            Weblog = '(btnsend--Access)'
            
        except selenium.common.exceptions.TimeoutException:
            driver.switch_to.frame('untie')
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "btnsend")))
            driver.switch_to.default_content()
            Weblog = '(btnsend switch frame--Access)'
        
        except:
            Weblog = '(btnsend--Fail)'
            
        app_log = app_log + Weblog
        
        if '(btnsend--Access)' in app_log:
            driver.find_element_by_name('btnsend').click()
        
        elif '(btnsend switch frame--Access)' in app_log:
            driver.switch_to.frame('untie')
            driver.find_element_by_name('btnsend').click()
            driver.switch_to.default_content()
        else:
            #driver.refresh()
            continue
        
        # 申請失敗訊息判斷
        try:
            WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "ErrorMsgText")))
            Weblog = '(ErrorMsgText--Access)'
        
        except:
            Weblog = '(ErrorMsgText--Notexist)'
            
        app_log = app_log + Weblog        
        
        if '(ErrorMsgText--Access)' in app_log:
            app_NoEL = '('+driver.find_element_by_id('ErrorMsgText').text+')'
        
        elif '(ErrorMsgText switch frame--Access)' in app_log:
            driver.switch_to.frame('untie')
            app_NoEL = '('+driver.find_element_by_id('ErrorMsgText').text+')'
            driver.switch_to.default_content()
            
        else:
            app_NoEL = ''
        
        app_log = app_log + app_NoEL
        
        if app_NoEL != '':
            appstatus = 'Y'
            app_result = 'N'
            insertdate = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            appdocs = (eaid,pid,addressi,landno,purpose,appstatus,entitytype,el_number,el_appdttm,app_result,insertdate,app_log)
            EL04_tmp_result = EL04_tmp_etl(appdocs)
            toSQL(EL04_tmp_result, totb, server, database, username, password)
            
            appobs = APPOBS(server,username,password,database,totb,entitytype)
            if i == (obs - 1) and appobs < obs:
                status = 'Re-Exec'
            elif i == (obs - 1):
                status = 'Done'
            else:
                status = 'Continue'            
            updateMGM(server,username,password,database,MGtb,appobs,status,apply_dttm,entitytype,Machine,EntityPhase,todtltb)
            
        else:
            appstatus = 'Y'
            app_result = 'Y'
            insertdate = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            
            # 申請成功訊息判斷
            try:
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "ApplyForm")))
                Weblog = '(ApplyForm--Access)'
                
            except selenium.common.exceptions.TimeoutException:
                #driver.switch_to.frame('untie')
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "ApplyForm")))
                driver.switch_to.default_content()
                Weblog = '(ApplyForm switch frame--Access)'
                
            except:
                Weblog = '(ApplyForm--Fail)'
            
            app_log = app_log + Weblog
            
            if '(ApplyForm--Access)' in app_log:
                app_EL = driver.find_element_by_id('ApplyForm').text
                el_number = app_EL[app_EL.find('\n')+1:app_EL.find('時間')]
                el_appdttm = app_EL[app_EL.find('時間:')+3:app_EL.find('\n申請作業成功')]
            
            if '(ApplyForm switch frame--Access)' in app_log:
                driver.switch_to.frame('untie')
                app_EL = driver.find_element_by_id('ApplyForm').text
                el_number = app_EL[app_EL.find('\n')+1:app_EL.find('時間')]
                el_appdttm = app_EL[app_EL.find('時間:')+3:app_EL.find('\n申請作業成功')]
                driver.switch_to.default_content()
                
            # 申請成功訊息判斷
            try:
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'Submit')))
                Weblog = '(Submit--Access)'
                
            except selenium.common.exceptions.TimeoutException:
                driver.switch_to.frame('untie')
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "Submit")))
                driver.switch_to.default_content()
                Weblog = '(Submit switch frame--Access)'
                
            except:
                Weblog = '(Submit--Fail)'
            
            app_log = app_log + Weblog
            
            if '(Submit--Access)' in app_log:
                driver.find_element_by_name('Submit').click()
            
            if '(Submit switch frame--Access)' in app_log:
                driver.switch_to.frame('untie')
                driver.find_element_by_name('Submit').click()
                driver.switch_to.default_content()
              
            appdocs = (eaid,pid,addressi,landno,purpose,appstatus,entitytype,el_number,el_appdttm,app_result,insertdate,app_log)
            EL04_tmp_result = EL04_tmp_etl(appdocs)
            toSQL(EL04_tmp_result, totb, server, database, username, password)
            appobs = APPOBS(server,username,password,database,totb,entitytype)
            appsuccessobs = APPSUCCESSOBS(server,username,password,database,fromtb,totb,entitytype)

            if i == (obs - 1) and appobs < obs:
                status = 'Re-Exec'
            elif i == (obs - 1):
                status = 'Done'
            else:
                status = 'Continue'
                
            if status == 'Done':
                _status = 'Stand by'
            else:
                _status = status
          
            updateMGM(server,username,password,database,MGtb,appobs,status,apply_dttm,entitytype,Machine,EntityPhase,todtltb)
            insertMG(server,username,password,database,MGtb,entitytype,Machine,EntityPhase_next,totb,appsuccessobs,_status,apply_dttm,EntityPath_next)
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME,'Image3')))
        driver.find_element(By.NAME, 'Image3').click()
        #driver.refresh()
    
    else:
        appdocs = (eaid,pid,addressi,landno,purpose,appstatus,entitytype,el_number,el_appdttm,app_result,insertdate,'Landno 不合規')
        EL04_tmp_result = EL04_tmp_etl(appdocs)
        toSQL(EL04_tmp_result, totb, server, database, username, password)
        appobs = APPOBS(server,username,password,database,totb,entitytype)
        if i == (obs - 1) and appobs < obs:
            status = 'Re-Exec'
        elif i == (obs - 1):
            status = 'Done'
        else:
            status = 'Continue'            
        updateMGM(server,username,password,database,MGtb,appobs,status,apply_dttm,entitytype,Machine,EntityPhase,todtltb)
        #driver.refresh()
    print(appobs)
    print(entitytype)
try:
    driver.close()
except:
    pass
import pymssql
conn = pymssql.connect(server=server, user=username, password=password, database = database)
cursor = conn.cursor()
        
script = f"""
    insert into EL04
    select * from {totb} where [eaid]  not in (select [eaid] from EL04 ) 
    """
cursor.execute(script)
conn.commit()
cursor.close()
conn.close()

obs = src_obs(server,username,password,database,fromtb,totb,entitytype)
print(obs)
if obs == 0 :
    SUCC_mail(obs)
    try:
        driver.close()
    except:
        pass
    os.system(rf"C:\Py_Project\project\pybatch\03.pybatch_EL_revproc.bat")
    
    pass
else:
    ERR_mail(obs)
    try:
        driver.close()
    except:
        pass
    os.system(rf"C:\Py_Project\project\pybatch\02.pybatch_EL_appproc.bat")
   
    pass