### 套件函式設定
import io
import csv
import time
import pandas as pd
from PIL import Image
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# CSV追加寫入函式
def AWCsv(file_path,datalist):
    csvFileToWrite = open(file_path, 'a', encoding='utf-8-sig', newline='')
    csvDataToWrite = csv.writer(csvFileToWrite)
    csvDataToWrite.writerow(datalist)
    csvFileToWrite.close()

# CSV覆蓋寫入函式(紀錄執行情況專用)
def OWcsv(file_path,exec_name,n):
    f = open(file_path,'w',encoding='utf8',newline='')
    with f:
        fcolumn = ['exec_name','exec_obs']
        writer = csv.DictWriter(f, fieldnames = fcolumn)
        writer.writeheader()
        writer.writerow({'exec_name' : exec_name, 'exec_obs': n})
    f.close()

# CSV偵測函式(紀錄執行情況專用)
def existfile(file_path):
    import os
    # 檢查檔案是否存在
    if os.path.isfile(file_path):
        return True
    else:
        return False

# CSV覆蓋讀取函式(紀錄執行情況專用)
def traceObs(file_path):
    trace = pd.read_csv(file_path).drop_duplicates()
    return trace['exec_obs'][0]

# Unicode數字辨識函式
def is_number(uchar):
    if uchar >= u'\u0030' and uchar <= u'\u0039':
        return True
    else:
        return False

# 確認網頁元素(ID)是否存在，存在返回flag=true，否則返回false
def isElementExist(element):
    flag=True
    try:
        driver.find_element_by_id(element)
        return flag
    except:
        flag=False
        return flag

# 確認模擬過程中，頁面是否自動跳轉回首頁，存在返回flag=true，否則返回false
def check_interface(element):
    driver.switch_to.default_content()
    home_page=True
    try:
        driver.find_element_by_id(element)
        return home_page
    except:
        flag=False
        return home_page

# 確認地段是否在列表中，存在返回true，否則返回false
def session_exist(char):
    if char in session_list:
        return True
    else:
        return False
    
    
    
### 參數設定

# 中華電信帳密
acc_name = '89850388'
acc_pawd = 'bvi0001'

# 網站URL
url = 'http://210.71.181.102/Home/Index'
Drvfile = r'C:\Py_Project\env\chromedriver_win32\chromedriver'

# 設定電謄文件類型
purpose = 'EL'

# job name
exec_name = rf'MG_P2_{purpose}'

# job status rec
trace_file = rf"C:\Py_Project\output\{exec_name}_TRACE.csv"

# 迴圈起始值設定
if existfile(trace_file):
    i = traceObs(trace_file) + 1
else:
    i = 0
    
# 目的地檔案偵測 + 寫入表頭
ta_file = rf"C:\Py_Project\output\{exec_name}_RESULT.csv"

if existfile(ta_file):
    pass
else:
    data = ('i','ID','casei',purpose,'apply_dttm')
    AWCsv(ta_file,data)

### 地號來源導入
path = r"C:\Py_Project\output\\"
comnm = 'LandNo.csv'
land_info = pd.read_csv(path+comnm).drop_duplicates().query("ID != 'ID'")
land_info = land_info.reset_index(drop=True)

### Connect to Website
driver = webdriver.Chrome(Drvfile)
driver.get(url)

# 勾選同意項目
checkboxs=driver.find_elements_by_css_selector('input[type=checkbox]')
checkboxs[0].click()
driver.find_element(By.ID, 'yes').click()

# 跳轉frame
driver.switch_to_frame('untie')
login_obj = driver.find_element(By.XPATH,'/html/body/table[2]/tbody/tr/td[1]/table/tbody/tr[1]/td').click()

try:
    element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "aa-uid")))
    driver.find_element(By.ID, 'aa-uid').send_keys(acc_name)
    driver.find_element(By.ID, 'aa-passwd').send_keys(acc_pawd)
finally:
    pass

img_ele = driver.find_element_by_id('AAAIden1')
img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
img

#輸入captcha文字
captchaID = input('')
driver.find_element_by_name('aa-captchaID').send_keys(captchaID)

#登入
driver.find_element_by_xpath('//*[@id="submit_hn"]/a').click()

for i in range(i-1,len(land_info)):
    #for i in range(40):

    time.sleep(1)
    driver.switch_to.default_content()
    driver.switch_to_frame('untie')
    driver.find_element_by_xpath('/html/body/table[2]/tbody/tr/td[1]/table/tbody/tr[2]/td/a').click()

    csaei = land_info['csaei'][i]
    ID = land_info['ID'][i]
    LandNo_f = land_info['LandNo_f'][i]
    LandNo_e = land_info['LandNo_e'][i]
    city = land_info['city'][i]
    town = land_info['town'][i]
    Lot = land_info['Lot'][i]

    for j in range(len(Lot)):
        if is_number(Lot[j]) == False:
            Lot_char = Lot[j:]
            Lot_char = Lot_char.replace(' ','')
            break

    #輸入城市
    s1 = Select(driver.find_element_by_id('City_ID'))
    s1.select_by_visible_text(city)
    
    try:
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "area_id")))
        s1 = Select(driver.find_element_by_id('area_id'))
        s1.select_by_visible_text(town)
    finally:
        pass
        
    #輸入申請用途
    s1 = Select(driver.find_element_by_id('applyfor'))
    s1.select_by_visible_text('自行參考')

    #輸入ID
    driver.find_element_by_id('CONSIGNOR_ID').send_keys(ID)

    #輸入統一編號
    driver.find_element_by_id('INPUT_015').send_keys(ID)
    
    #輸入地段
    soup = BeautifulSoup(driver.page_source,'html.parser')
    session_list = []
    
    for session in soup.find_all(class_ = 'session_name'):
        session_list.append(session.text)
    
    if session_exist(Lot_char):
        driver.find_element(By.ID, "session_title").click()
        driver.find_element_by_name(Lot_char).click()
    else:
        tb = driver.find_element_by_xpath("//*[@id='table1']/tbody/tr[2]/td[8]/a")
        tb.click()
        continue

    # 輸入 'ID' + '地號'
    driver.find_element(By.ID,'INPUT_013').send_keys(LandNo_f+LandNo_e)
    
    # 勾選 EL & 異動索引    
    if purpose == 'EL':
        driver.find_element_by_name('INPUT_021').click()
        CsvToPath = rf'C:\Py_Project\output\MG_P2_EL_RESULT.csv'
    elif purpose == 'CNG':
        driver.find_element_by_name('INPUT_061').click()
        time.sleep(1)
        driver.find_element_by_xpath('/html/body/table[2]/tbody/tr/td/form[2]/div[2]/table[2]/tbody/tr[5]/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td[5]/span/input').click()
        time.sleep(1)
        CsvToPath = rf'C:\Py_Project\output\MG_P2_CNG_RESULT.csv'
    else:
        print('please set purpose correctly!')
        break

    #勾選新增資料    
    driver.find_element_by_name('btnnew').click()
    time.sleep(1)

    #勾選送出
    driver.find_element_by_name('btnsend').click()
    OWcsv(trace_file,exec_name,i)
    time.sleep(2)

    if isElementExist("ErrorMsgText"):
        print(i,ID,csaei,'without information','without information')
        tb = driver.find_element_by_xpath("//*[@id='table1']/tbody/tr[2]/td[8]/a")
        tb.click()
    else:
        soup = BeautifulSoup(driver.page_source,'html.parser')
        output = soup.find_all('b')[1].text
        El = output[0:output.find(' ')]
        datetime = output[output.find(':')+1:].strip()

        driver.find_element_by_name('Submit').click()
        time.sleep(1)
        driver.switch_to.default_content()
        driver.switch_to_frame('untie')
        print(i,ID,csaei,purpose,El,datetime)
        data = (i,ID,csaei,purpose,El,datetime)
        AWCsv(CsvToPath,data)
        tb = driver.find_element_by_xpath("//*[@id='table1']/tbody/tr[2]/td[8]/a")
        tb.click()

    if i == len(land_info) - 1:
        print('Done')