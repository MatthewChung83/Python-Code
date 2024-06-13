# 匯入外部python庫
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium import webdriver
from bs4 import BeautifulSoup
from urllib.parse import unquote

from PIL import Image
#import import_ipynb
import pyautogui
import time
import re
import pdfkit

# 自定義python庫
from dict import *
from config import *
from etl_func import *

# 系統相關路徑資訊
server,database,username,password,totb1,totb2 = db['server'],db['database'],db['username'],db['password'],db['totb1'],db['totb2']
print_path,captcha_path = captchainfo['print_path'],captchainfo['captcha_path']

# 簡易案件 & 裁判書 等兩種查詢方式
try:
    for ur in range(len(urls)):
    
        Drvfile = wbinfo['Drvfile']
        driver = webdriver.Chrome(Drvfile)
        driver.get(urls[ur][0])
        time.sleep(3)
        driver.maximize_window()
        defense_anti(driver)
        findelement_LINK_click(driver,"行動版")
        
        # 拍賣,代位,分割,抵押權,變價,塗銷,撤銷 等7種 關鍵字搜尋
        for qtype in qtypes:
            findelement_CSS_click(driver,".fa-3x")
            time.sleep(1)
            findelement_LINK_click(driver,urls[ur][1])
            findelement_ID_click(driver,"txtKW")
            findelement_ID_sendkeys(driver,"txtKW",qtype)
            findelement_ID_click(driver,"btnSubmit")
            time.sleep(1)
            
            # 遍掃返回結果頁的頁數
            for p in range(25):
                soup = BeautifulSoup(driver.page_source,'lxml')
                
                # 頁面資訊列表解析
                for c in soup.find_all(id = 'hlTitle'):
                    t = c.find_parents("td")[0].text.replace('\n',',').replace('\r',',').replace(' ','').split(',')
                    tlist = list(filter(None,t))
                    tlist.append(tlist[2][-9:])
                    print(tlist)
                    try :
                        if tlist[4] not in calendar(str_date,end_date):
                            break
                    except:
                        if tlist[3] not in calendar(str_date,end_date):
                            break
                    tlist[2] = tlist[2].replace(tlist[4],'')
                    # 法文解析            
                    findelement_XPATH_href(driver,c['href'])
                    soup_c = BeautifulSoup(driver.page_source,'lxml')            
                    parse_result = parse(soup_c,'htmlcontent','text-pre text-pre-in',splits,deletes)
                    content = parse_result[0]
                    s_parse = parse_result[1]
                    e_parse = parse_result[2]
                    main_sta = parse_result[3]           
                    roles_names = roles_process(roles,main_sta)
                    insertdate = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))            
                    tlist = tlist + ['https://law.judicial.gov.tw/LAW_MOBILE/FJUD/'+c['href'],qtype,s_parse,e_parse,main_sta]
                    
                    # 法文內拆解每個角色，辨識姓名
                    for rn in range(len(roles_names)):
                        tlist.append(roles_names[rn][1][0:3500])
                        name = spname(delimis,roles_names[rn][1])
                        print(name)
                        # 辨識後之姓名，入SQL DB：FJUD_Default_AD_detail
                        if len(name) > 0 and src_obs(server,username,password,database,'FJUD_Default_AD_detail','link',tlist[5]) == 0:
                            docs = ['https://law.judicial.gov.tw/LAW_MOBILE/FJUD/'+c['href'],roles_en[rn],name,insertdate]
                            etls = indetail(docs)
                            toSQL(etls, totb2, server, database, username, password)
                    
                    article = content[:3500]
                    tlist = tlist + [article,insertdate]
                    print(tlist)
                    
                    
                    # 完整法文文字檔，入SQL DB：FJUD_Default_AD
                    if src_obs(server,username,password,database,'FJUD_Default_AD_detail','link',tlist[5]) == 0:
                        url1 = tlist[5]
                        html = unquote(url1).replace(',','%2c').replace('/LAW_MOBILE','')+'&ot=in'
                        config = pdfkit.configuration(wkhtmltopdf="C:\Program Files\wkhtmltopdf\wkhtmltopdf.exe")
                        filename =  rf'C:\Py_Project\project\FJUD_Default_AD\file\\'+tlist[1]+'_'+tlist[2]+'_'+tlist[4]+'_'+tlist[6]+'.pdf'
                        pdfkit.from_url(html,filename,configuration=config)
                        doc = [tlist[0],tlist[1],tlist[2],tlist[3],tlist[4],tlist[5],tlist[6],tlist[7],tlist[8],tlist[9],tlist[10],tlist[11],tlist[12],tlist[13],tlist[14],tlist[15],tlist[16],tlist[17],tlist[18],filename]
                        etl = indata(doc)
                        toSQL(etl, totb1, server, database, username, password)
                    
                    driver.back()
                try:
                    if tlist[4] not in calendar(str_date,end_date):
                        break
                except:
                    if tlist[3] not in calendar(str_date,end_date):
                        break
                
                findelement_ID_click(driver,"hlNext")
                time.sleep(3)
        
        driver.quit()
except:
    pass
import shutil
import os
import time
import datetime


file_source = rf'C:\\Py_Project\\project\\FJUD_Default_AD\\file\\'
file_destination = rf'\\fortune\Cashfile\UCS\FJUD_Default_AD\\'
 
get_files = os.listdir(file_source)

time.strftime('%Y%m%d')
now_time = datetime.datetime.now()
future_time = now_time + datetime.timedelta(days=-4)
fu = future_time.strftime('%Y%m%d')
print(fu)

for g in get_files:
    t = time.localtime(os.path.getctime(file_source+g))
    ctime=time.strftime("%Y%m%d",t)
    
    if ctime > fu :     
        shutil.copy(file_source + g, file_destination)
        print(ctime)
        print(g)
        pass
    
    else :

        pass
#Send Mail        
import smtplib
from email.mime.text import MIMEText
from email.header import Header
sender = 'matthew5043@ucs.com'
receivers = ['matthew5043@ucs.com'] # 接收郵件
# 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
message = MIMEText('新資源FJUD_Default_AD上載圖檔已完成', 'plain', 'utf-8')
message['From'] = Header('wbt', 'utf-8') # 傳送者
message['To'] = Header('matthew5043', 'utf-8') # 接收者
subject = '新資源FJUD_Default_AD上載圖檔已完成'
message['Subject'] = Header(subject, 'utf-8')
try:
    smtpObj = smtplib.SMTP('vrh19.ucs.com')
    smtpObj.sendmail(sender, receivers, message.as_string())
    print('郵件傳送成功')
except smtplib.SMTPException:
    print('Error: 無法傳送郵件')