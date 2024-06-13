import io
import time
import calendar
import requests
import pyautogui
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

import import_ipynb

from config import *
from etl_func import *
from batch_args import *
from OCR_Mod import *

acc_name,acc_pwd,url,Drvfile = wbinfo['acc_name'],wbinfo['acc_pwd'],wbinfo['url'],wbinfo['Drvfile']
Apurl,imgf,imgp = APinfo['Apurl'],APinfo['imgf'],APinfo['imgp']

driver = webdriver.Chrome(Drvfile)
driver.get(url)

revlog = ''

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
dlist_sort[6]

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


soup = BeautifulSoup(driver.page_source,'html.parser')
td_text = soup.find_all(id = 'form1')[0].find_all('tr')
trset = []

for tr in range(len(td_text)):
    tdset = []
    
    for td in range(len(td_text[tr].find_all('font'))):
        tdset.append(td_text[tr].find_all('font')[td].text.replace('\n','').strip())
        
        if td_text[tr].find_all('font')[td].find('a'):
            tdset.append(td_text[tr].find_all('font')[td].find('a'))
                
    print(tdset)
        
    if tdset[0] == 'GUID':
        continue
        
    else:
        print(tdset)
        ### 下載PDF 模擬鍵盤操作 ###
        
        ### 依每個文件對應的 JS參數 點開特定連結
        livalue = 'a[onclick*=' + tdset[7].attrs['onclick'] + ']'
        
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
            
        #關閉下載視窗
        time.sleep(0.5)
        driver.close()
        
        # 切回主頁面
        driver.switch_to.window(windows[0])
        driver.switch_to.frame('untie')
        time.sleep(0.5)
        
driver.find_element_by_xpath ("//li[contains( text( ),'下一頁')]").click()