import io
import time
import calendar
import requests
import pyautogui
from PIL import Image
from bs4 import BeautifulSoup
from selenium import webdriver
import selenium.common.exceptions
import os
import datetime
from datetime import timedelta
import shutil
import os

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

# Changes Windows 1920*1080
import win32api
dm = win32api.EnumDisplaySettings(None, 0)
dm.PelsHeight = 1080
dm.PelsWidth = 1920
dm.BitsPerPel = 32
dm.DisplayFixedOutput = 0
win32api.ChangeDisplaySettings(dm, 0)


Machine,server,database,username,password = hostname(),db['server'],db['database'],db['username'],db['password']
servicetb,fromtb,totb,entitytype = db['servicetb'],db['fromtb'],db['totb'],db['entitytype']
MGtb,fromtb,totb,EntityPhase,EntityPhase_last,EntityPath,EntityPath_last = db['MGtb'],db['fromtb'],db['totb'],db['EntityPhase'],db['EntityPhase_last'],db['EntityPath'],db['EntityPath_last']
acc_name,acc_pwd,url,Drvfile = wbinfo['acc_name'],wbinfo['acc_pwd'],wbinfo['url'],wbinfo['Drvfile']
Apurl,imgf,imgp = APinfo['Apurl'],APinfo['imgf'],APinfo['imgp']

driver = webdriver.Chrome(Drvfile)
driver.get(url)

#createtmp(server,username,password,database,fromtb,totb,entitytype)
obs = src_obs(server,username,password,database,fromtb,totb,entitytype)
revlog = ''


now = datetime.datetime.now()
this_week_start = now - timedelta(days=now.weekday())
FD=this_week_start.strftime('%Y%m%d')
weekly='AM_Weekly_89850389_'+FD+'_EL'
import os
import datetime
from datetime import timedelta
today = datetime.datetime.now().strftime('%Y%m%d')
year = datetime.datetime.now().strftime('%Y')

try:
    os.makedirs(rf'C:\Py_Project\project\easymap_EL_1.2\gen_el05\file\{weekly}')
except:
    pass
try:
    os.makedirs(rf'\\vnt07\ucsfile\el\{weekly}')
except:
    pass
try:
    os.makedirs(rf'\\fortune\CashFile\{year}\EL\{today}')
except:
    pass
try:
    os.makedirs(rf'//fortune/Cashfile/UCS/EL/{weekly}/')
except:
    pass
    

### [進入網站] 勾選同意 ###
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'ok')))
    Weblog = '(ok--Access)'
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'ok')))
    driver.switch_to.default_content()
    Weblog = '(ok switch frame--Access)'
except:
    Weblog = '(ok--Fail)'
    
revlog = revlog + Weblog

if '(ok--Access)' in revlog:
    driver.find_element_by_id('ok').click()
    
elif '(ok switch frame--Access)' in revlog:
    driver.switch_to.frame('untie')
    driver.find_element(By.ID,'ok').click()
    driver.switch_to.default_content()
    
else:
    driver.quit()

### [進入網站] 同意進入 ###
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'yes')))
    Weblog = '(yes--Access)'
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'yes')))
    driver.switch_to.default_content()
    Weblog = '(yes switch frame--Access)'
except:
    Weblog = '(yes--Fail)'
    
revlog = revlog + Weblog

if '(yes--Access)' in revlog:
    driver.find_element_by_id('yes').click()
    
elif '(yes switch frame--Access)' in revlog:
    driver.switch_to.frame('untie')
    driver.find_element_by_id('yes').click()
    driver.switch_to.default_content()
    
else:
    driver.quit()

### [進入網站] 登入帳號 ###
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'Img2')))
    Weblog = '(Img2--Access)'
    
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'Img2')))
    driver.switch_to.default_content()
    Weblog = '(Img2 switch frame--Access)'
    
except:
    Weblog = '(Img2--Fail)'
    
revlog = revlog + Weblog

if '(Img2--Access)' in revlog:
    driver.find_element_by_id('Img2').click()
    
elif '(Img2 switch frame--Access)' in revlog:
    driver.switch_to.frame('untie')
    driver.find_element_by_id('Img2').click()
    driver.switch_to.default_content()
    
else:
    driver.quit()

q = 0   # 索引變數
while q < 30:    
    ### [進入網站] 輸入帳號 ###
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'aa-uid')))
        Weblog = '(aa-uid--Access)'
        
    except selenium.common.exceptions.TimeoutException:
        driver.switch_to.frame('untie')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'aa-uid')))
        driver.switch_to.default_content()
        Weblog = '(aa-uid switch frame--Access)'
        
    except:
        Weblog = '(aa-uid--Fail)'
        
    revlog = revlog + Weblog
    
    if '(aa-uid--Access)' in revlog:
        driver.find_element(By.ID, 'aa-uid').send_keys(acc_name)
        
    elif '(aa-uid switch frame--Access)' in revlog:
        driver.switch_to.frame('untie')
        driver.find_element(By.ID, 'aa-uid').send_keys(acc_name)
        driver.switch_to.default_content()
    
    else:
        driver.quit()
    
    ### [進入網站] 輸入密碼 ###
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'aa-passwd')))
        Weblog = '(aa-passwd--Access)'
    except selenium.common.exceptions.TimeoutException:
        driver.switch_to.frame('untie')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'aa-passwd')))
        driver.switch_to.default_content()
        Weblog = '(aa-passwd switch frame--Access)'
    except:
        Weblog = '(aa-passwd--Fail)'
        
    revlog = revlog + Weblog
    
    if '(aa-passwd--Access)' in revlog:
        driver.find_element(By.ID, 'aa-passwd').send_keys(acc_pwd)
        
    elif '(aa-passwd switch frame--Access)' in revlog:
        driver.switch_to.frame('untie')
        driver.find_element(By.ID, 'aa-passwd').send_keys(acc_pwd)
        driver.switch_to.default_content()
        
    else:
        driver.quit()
    
    ### [進入網站] 解析captcha ###
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'AAAIden1')))
        Weblog = '(AAAIden1--Access)'
    except selenium.common.exceptions.TimeoutException:
        driver.switch_to.frame('untie')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'AAAIden1')))
        driver.switch_to.default_content()
        Weblog = '(AAAIden1 switch frame--Access)'
    except:
        Weblog = '(AAAIden1--Fail)'
        
    revlog = revlog + Weblog
    
    if '(AAAIden1--Access)' in revlog:
        img_ele = driver.find_element_by_id('AAAIden1')
        img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
        img.save(imgp)
        
    elif '(AAAIden1 switch frame--Access)' in revlog:
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
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME,'aa-captchaID')))
        Weblog = '(aa-captchaID--Access)'
    except selenium.common.exceptions.TimeoutException:
        driver.switch_to.frame('untie')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME,'aa-captchaID')))
        driver.switch_to.default_content()
        Weblog = '(aa-captchaID switch frame--Access)'
    except:
        Weblog = '(aa-captchaID--Fail)'
    
    revlog = revlog + Weblog
    
    if '(aa-captchaID--Access)' in revlog:
        driver.find_element_by_name('aa-captchaID').send_keys(captchaocr)
        time.sleep(2)
    
    elif '(aa-captchaID switch frame--Access)' in revlog:
        driver.switch_to.frame('untie')
        driver.find_element_by_name('aa-captchaID').send_keys(captchaocr)
        time.sleep(2)
        driver.switch_to.default_content()
        
    else:
        driver.quit()        
    
    ### [進入網站] 進入網站 ###
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'submit_hn')))
        Weblog = '(submit_hn--Access)'
        
    except selenium.common.exceptions.TimeoutException:
        driver.switch_to.frame('untie')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'submit_hn')))
        driver.switch_to.default_content()
        Weblog = '(submit_hn switch frame--Access)'
        
    except:
        Weblog = '(submit_hn--Fail)'
    
    revlog = revlog + Weblog
    
    if '(submit_hn--Access)' in revlog:
        driver.find_element(By.ID, 'submit_hn').click()
    
    elif '(submit_hn switch frame--Access)' in revlog:
        driver.switch_to.frame('untie')
        driver.find_element(By.ID, 'submit_hn').click()
        driver.switch_to.default_content()
        
    else:
        driver.quit()        
    
    try:
        WebDriverWait(driver, 3).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        time.sleep(2)
        alert.accept()
        time.sleep(2)
        print("alert accepted")
    
    except TimeoutException:
        print("no alert")
        break
        
    q += 1

entrylog = ''
### [進入網站] 進入領件區 ###
try:
    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME,'Image15')))
    Weblog = '(Image15--Access)'
    
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME,'Image15')))
    driver.switch_to.default_content()
    Weblog = '(Image15 switch frame--Access)'
    
except:
    Weblog = '(Image15--Fail)'

revlog = revlog + Weblog

if '(Image15--Access)' in revlog:
    driver.find_element_by_name('Image15').click()

elif '(Image15 switch frame--Access)' in revlog:
    driver.switch_to.frame('untie')
    driver.find_element_by_name('Image15').click()
    driver.switch_to.default_content()
    
else:
    driver.quit()

rec_log = ''

# 選擇領件日期        
try:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'DateFrom')))
    driver.switch_to.default_content()
    Weblog = '(DateFrom switch frame--Access)'
    
except selenium.common.exceptions.TimeoutException:
    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'DateFrom')))
    Weblog = '(DateFrom--Access)'
    
except:
    Weblog = '(DateFrom--Fail)'
    
revlog = revlog + Weblog

if '(DateFrom--Access)' in revlog:
    pass

elif '(DateFrom switch frame--Access)' in revlog:
    driver.switch_to.frame('untie')
    
else:
    driver.refresh()

soup = BeautifulSoup(driver.page_source,'html.parser')

dlist = []
for d in soup.find_all(id = 'DateFrom')[0].find_all('option'):
    dlist.append(d.text)
    
dlist_sort = sorted(dlist, reverse = False)
dlist_sort[0]

seldate = Select(driver.find_element_by_id('DateFrom'))
seldate.select_by_visible_text(dlist_sort[0])

# 選擇文件狀態       
try:
    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'Status')))
    Weblog = '(Status--Access)'
    
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'Status')))
    driver.switch_to.default_content()
    Weblog = '(Status switch frame--Access)'
    
except:
    Weblog = '(Status--Fail)'
    
revlog = revlog + Weblog

if '(Status--Access)' in revlog:
    selstatus = Select(driver.find_element_by_id('Status'))
    #selstatus.select_by_visible_text('未下載')
    selstatus.select_by_visible_text('全部')
    
elif '(Status switch frame--Access)' in revlog:
    driver.switch_to.frame('untie')
    selstatus = Select(driver.find_element_by_id('Status'))
    #selstatus.select_by_visible_text('未下載')
    selstatus.select_by_visible_text('全部')
    driver.switch_to.default_content()
    
else:
    driver.refresh()
    
# 送出查詢       
try:
    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME,'Submit')))
    Weblog = '(Submit--Access)'
    
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME,'Submit')))
    driver.switch_to.default_content()
    Weblog = '(Submit switch frame--Access)'
    
except:
    Weblog = '(Submit--Fail)'
    
revlog = revlog + Weblog

if '(Submit--Access)' in revlog:
    driver.find_element_by_name('Submit').click()
    
elif '(Submit switch frame--Access)' in revlog:
    driver.switch_to.frame('untie')
    driver.find_element_by_name('Submit').click()
    driver.switch_to.default_content()
    
else:
    driver.refresh()

soup = BeautifulSoup(driver.page_source,'html.parser')
pset = []
for p in soup.find_all(id = 'pager')[0].find_all('li'):
    if '第一頁' in p.text or '上一頁' in p.text or '下一頁' in p.text or '最後頁' in p.text:
        pass
    else:
        pset.append(int(p.text))
pset = sorted(pset, reverse = True)


#for i in range(1):
for i in foo(-1,obs-1):
    driver.find_element_by_xpath ("//li[contains( text( ),'1')]").click()
    time.sleep(1)
    #rev_info = dbfrom(server,username,password,database,fromtb,totb,entitytype)[0]
    rev_info = dbfrom(server,username,password,database,fromtb,totb,entitytype)[0]
    eaid,pid,addressi,landno = rev_info[0],rev_info[1],rev_info[2],rev_info[3]
    purpose,el_number,el_appdttm = rev_info[4],rev_info[6].strip(),rev_info[7]
    appdt = el_appdttm.strip().replace('/','').replace(':','').replace(' ','_')    
    
    rev_appdttm = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    insertdate = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    savename = f'{eaid}_{addressi}_{entitytype}'
    
    ta_file_vnt07 = rf'\\vnt07\ucsfile\el\{weekly}\{entitytype}_{appdt}_{eaid}_{addressi}.pdf'
    ta_file = rf'C:\Py_Project\project\easymap_EL_1.2\gen_el05\file\{weekly}\{entitytype}_{appdt}_{eaid}_{addressi}.pdf'
    #ta_file = '\\vnt07\ucsfile\el\'+entitytype+'\{entitytype}_{appdt}_{eaid}_{addressi}.pdf'
    
    number_src,file_save,el_appdttm,note,pages,app_fee = '','','','','',''
    
    if existfile(ta_file):
        note = 'been downloaded in target file'
        docs = (eaid,pid,addressi,landno,purpose,entitytype,number_src,el_number,pages,app_fee,ta_file_vnt07,el_appdttm,insertdate,note,insertdate,revlog)
        EL05_tmp_result = EL05_tmp_etl(docs)
        print(EL05_tmp_result)
        toSQL(EL05_tmp_result, totb, server, database, username, password)                
        revobs = REVOBS(server,username,password,database,fromtb,totb,entitytype)
        appobs = APPOBS(server,username,password,database,fromtb,entitytype)
                
        if revobs < appobs:
            status = 'Continue'
        else:
            status = 'Done'

        updateMGM(server,username,password,database,MGtb,revobs,status,rev_appdttm,entitytype,Machine,EntityPhase_last,fromtb)
        continue

    for p in range(pset[0]):
        
        soup = BeautifulSoup(driver.page_source,'html.parser')
        #td_text = soup.find_all('tr',class_ = 'bgy')
        td_text = soup.find_all(id = 'form1')[0].find_all('tr')
        trset = []
        for tr in range(len(td_text)):
            tdset = []
            
            for td in range(len(td_text[tr].find_all('font'))):
                tdset.append(td_text[tr].find_all('font')[td].text.replace('\n','').strip())
                
                if td_text[tr].find_all('font')[td].find('a'):
                    tdset.append(td_text[tr].find_all('font')[td].find('a'))
           
            if tdset[2] == el_number:
                trset.append(tdset)
                el_appdttm,number_src,el_number,purpose,pages,app_fee,doc_href = trset[0][0],trset[0][1],trset[0][2],trset[0][3],trset[0][4],trset[0][5],trset[0][7]

            ### 下載PDF 模擬鍵盤操作 ###
                
                ### 依每個文件對應的 JS參數 點開特定連結
                livalue = 'a[onclick*=' + trset[0][7].attrs['onclick'] + ']'
                
                ### 區分已下載與未下載的文件連結
                if 'LoadPackage6' in livalue:
                    ckvalue = livalue[livalue.find("'LoadPackage6','','")+len("'LoadPackage6','','"):livalue[livalue.find("'LoadPackage6','','")+len("'LoadPackage6','','"):].find("'")+livalue.find("'LoadPackage6','','")+len("'LoadPackage6','','")]
                else:
                    ckvalue = livalue[livalue.find("'pdf','")+len("'pdf','"):livalue[livalue.find("'pdf','")+len("'pdf','"):].find("'")+livalue.find("'pdf','")+len("'pdf','")]
                driver.find_element_by_css_selector(f"a[onclick*='{ckvalue}']").click()
                
                # 獲取開啟的多個視窗控制代碼
                windows = driver.window_handles
                
                # 切換到當前最新開啟的視窗
                driver.switch_to.window(windows[-1])
                driver.maximize_window()
                
                # 點選確定付費下載
                if 'LoadPackage6' not in livalue:
                    driver.find_element_by_name('yes').click()
                    
                # 切換到當前最新開啟的視窗
                driver.switch_to.window(windows[-1])
                driver.maximize_window()
                time.sleep(1)
                               
                # 快捷鍵另存新檔
                #pyautogui.click(1500,500,button='right') 
                pyautogui.rightClick(1500,500)
                pyautogui.press('a')
                time.sleep(2)
                
                # 切換英文輸入法
                #pyautogui.press('shift')
                time.sleep(2)
                
                # 儲存檔案
                pyautogui.typewrite(ta_file)
                time.sleep(2)
                pyautogui.press('enter')
                time.sleep(2)
                #關閉下載視窗
                driver.close()
                
                # 切回主頁面
                driver.switch_to.window(windows[0])
                driver.switch_to.frame('untie')
                time.sleep(2)
                
            ### 下載PDF 模擬鍵盤操作 ###                
                
                docs = (eaid,pid,addressi,landno,purpose,entitytype,number_src,el_number,pages,app_fee,ta_file_vnt07,el_appdttm,rev_appdttm,note,insertdate,revlog)
                EL05_tmp_result = EL05_tmp_etl(docs)
                print(EL05_tmp_result)
                toSQL(EL05_tmp_result, totb, server, database, username, password)                
                revobs = REVOBS(server,username,password,database,fromtb,totb,entitytype)
                appobs = APPOBS(server,username,password,database,fromtb,entitytype)
                
                
                if revobs < appobs:
                    status = 'Continue'
                else:
                    status = 'Done'

                updateMGM(server,username,password,database,MGtb,revobs,status,rev_appdttm,entitytype,Machine,EntityPhase_last,fromtb)
            else:
                pass
        try:
            driver.find_element_by_xpath ("//li[contains( text( ),'下一頁')]").click()
        except:
            pass

try:
    driver.close()
except:
    pass
import pymssql
conn = pymssql.connect(server=server, user=username, password=password, database = database)
cursor = conn.cursor()
        
script = f"""
    insert into EL05
    select * from {totb} where [eaid]  not in (select [eaid] from EL05 ) 
    """
cursor.execute(script)
conn.commit()
cursor.close()
conn.close()

import pymssql
conn = pymssql.connect(server=server, user=username, password=password, database = database)
cursor = conn.cursor()
        
script = f"""
    select distinct
    a.*,
    case b.flg  
    when 'same as landno' then '戶政'
    when 'new by buildno' then '地政'
    end as flg
    into #123
    from EL05 a ,EL03DTL b where a.eaid = b.eaid and a.LandNo = b.LandNo


     update EL05
     set note = (select flg from #123 where EL05.eaid = #123.eaid and EL05.LandNo =#123.LandNo)

    """
cursor.execute(script)
conn.commit()
cursor.close()
conn.close()
try:
    driver.close()
except:
    pass
pass
#os.makedirs(rf'C:\Py_Project\project\easymap_EL_1.2\gen_el05\file\{weekly}')
#os.makedirs(rf'\\vnt07\ucsfile\el\{weekly}')
#os.makedirs(rf'//fortune/Cashfile/UCS/EL/{weekly}/')
obs = src_obs(server,username,password,database,fromtb,totb,entitytype)
try:
 
    file_source = rf'C:\Py_Project\project\easymap_EL_1.2\gen_el05\file\{weekly}\\'
    file_destination = rf'\\vnt07\ucsfile\el\{weekly}\\'
 
    get_files = os.listdir(file_source)
 
    for g in get_files:
        shutil.copy(file_source + g, file_destination)
    
    file_source1 = rf'\\vnt07\ucsfile\el\{weekly}\\'
    file_destination1 = rf'//fortune/Cashfile/UCS/EL/{weekly}/'
 
    get_files1 = os.listdir(file_source1)
 
    for g in get_files1:
        shutil.copy(file_source1 + g, file_destination1)
    check = 'Y'
except:
    ERR_mail(obs)
    pass
obs = src_obs(server,username,password,database,fromtb,totb,entitytype)
print(obs)
if obs == 0 and check == 'Y':
    SUCC_mail(obs)
    try:
        driver.close()
    except:
        pass
    pass
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    insert into check_point_EL
    values('Y','{entitytype}')
    """    
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    os.system(rf"C:\Py_Project\project\pybatch\05.pdf_reader_el.bat")
    
else:
    ERR_mail_step2(obs)
    try:
        driver.close()
    except:
        pass
    pass