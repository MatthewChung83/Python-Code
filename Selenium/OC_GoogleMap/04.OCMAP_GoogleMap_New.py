# -*- coding: utf-8 -*-
"""
Created on Thu Feb  2 10:50:28 2023

@author: admin
"""









import io
import time 
import calendar
import requests
import pyautogui
import os

from PIL import Image
from bs4 import BeautifulSoup
from selenium import webdriver
import selenium.common.exceptions
import pandas as pd
import datetime

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.expected_conditions import alert_is_present

wbinfo = {
#'url':'https://accounts.google.com/v3/signin/identifier?dsh=S-905745455%3A1673592426937757&continue=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fmid%3D14aNasVz2eWaGT-LQlnHr0ISq65CIz1I%26usp%3Dsharing&ec=GAZA2gE&followup=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fmid%3D14aNasVz2eWaGT-LQlnHr0ISq65CIz1I%26usp%3Dsharing&passive=1209600&flowName=GlifWebSignIn&flowEntry=ServiceLogin&ifkv=AeAAQh5pMeQT6U3BkJYcyKkdTzCyCFDbNnktO-x1onChrkk_oVnBB9oHoGV5Kah_SwGCJ282vKd1VQ',
'Drvfile':rf'C:\Py_Project\env\chromedriver_win32\chromedriver',
#'All':rf'\\fortune\UCS\DI\OC\SCOTT4162\SCOTT4162_All.xlsx',
#'TSB':rf'\\fortune\UCS\DI\OC\SCOTT4162\SCOTT4162_TSB.xlsx',}
}
#url,Drvfile,All,TSB = wbinfo['url'],wbinfo['Drvfile'],wbinfo['All'],wbinfo['TSB']
Drvfile = wbinfo['Drvfile']

urll = [
        #'https://accounts.google.com/v3/signin/identifier?dsh=S-117041837%3A1675754453002914&continue=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fmid%3D1hmxln_fpFmOIOcM7Y7zupuVuXaa1lqg%26usp%3Dsharing&ec=GAZA2gE&followup=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fmid%3D1hmxln_fpFmOIOcM7Y7zupuVuXaa1lqg%26usp%3Dsharing&passive=1209600&flowName=GlifWebSignIn&flowEntry=ServiceLogin&ifkv=AWnogHdDbBKTojejGlAp3jKJS8LSsA1qpR9wDWIFiD7Zk0il1v48aRoAzffjDexSUK33EWjkkJNJ'+','+'ANDERSON',
       'https://accounts.google.com/v3/signin/identifier?dsh=S-1593203054%3A1681712270285015&continue=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1uYIgPtwrNX2S7zqsKhw6Ryj_3kjqWUw%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&ec=GAZA2gE&followup=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1uYIgPtwrNX2S7zqsKhw6Ryj_3kjqWUw%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&hl=zh-TW&ifkv=AQMjQ7T6aNN_tYqzul15RPFNc8W25AGqbOmKeKdf3L2cgEg413d7oqUAtB4PMa0BzyMTLa1iI1MRkQ&passive=1209600&flowName=GlifWebSignIn&flowEntry=ServiceLogin'+','+'ALEXY',
       'https://accounts.google.com/v3/signin/identifier?dsh=S655048351%3A1681711320362031&continue=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1qblXja7GTajvC4B282yFr8Xw9XPi26E%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&ec=GAZA2gE&followup=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1qblXja7GTajvC4B282yFr8Xw9XPi26E%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&hl=zh-TW&ifkv=AQMjQ7R_rfrdycglT3XenTu7AKFfPXLEiY2Bx2awTQB_-O7l0_1FOurkuZXk9BwT_q80EmtCQeUu1Q&passive=1209600&flowName=GlifWebSignIn&flowEntry=ServiceLogin'+','+'ANDERSON',
       'https://accounts.google.com/v3/signin/identifier?dsh=S-234206541%3A1681711357745882&continue=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1zhUvtlOXuzHbdAePLMLG6hcQsZTXWeg%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&ec=GAZA2gE&followup=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1zhUvtlOXuzHbdAePLMLG6hcQsZTXWeg%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&hl=zh-TW&ifkv=AQMjQ7QGizQqGu1ufOj2OkCJ8gzd7PCHiCKaYrhuY8_zgKpTZPO4LeHttY0SQvJ76I9_NRSIbQOfwQ&passive=1209600&flowName=GlifWebSignIn&flowEntry=ServiceLogin'+','+'BEN4423',
       'https://accounts.google.com/v3/signin/identifier?dsh=S1275321292%3A1681711409025665&continue=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1wrFPoRKpm2Ldj2hzrbIh3aGf3A2mwUU%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&ec=GAZA2gE&followup=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1wrFPoRKpm2Ldj2hzrbIh3aGf3A2mwUU%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&hl=zh-TW&ifkv=AQMjQ7TDXRLFPjFpl1wVS4eo79b22-W4jH3XS3DJr09TGjBFSt7OnnzbMbEuL7elvIyX-XhMZgE3&passive=1209600&flowName=GlifWebSignIn&flowEntry=ServiceLogin'+','+'KEVIN4584',
       'https://accounts.google.com/v3/signin/identifier?dsh=S1242336415%3A1681711438748969&continue=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1MQwYjQuvQ7qjl4VYym8_aTOl5APqBzA%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&ec=GAZA2gE&followup=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1MQwYjQuvQ7qjl4VYym8_aTOl5APqBzA%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&hl=zh-TW&ifkv=AQMjQ7QSFjDh2Q0o5AjbjzRRCWDVE7MQPjpxZeg1yPtLwj2jyUuAV0Dom5yFHWCIiaeY5zdN3JMixw&passive=1209600&flowName=GlifWebSignIn&flowEntry=ServiceLogin'+','+'JACKYH',
       'https://accounts.google.com/v3/signin/identifier?dsh=S386651260%3A1681711519884945&continue=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D11bK3P3DlFTIc3GQHyNY0Yo7Kf5e8fHk%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&ec=GAZA2gE&followup=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D11bK3P3DlFTIc3GQHyNY0Yo7Kf5e8fHk%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&hl=zh-TW&ifkv=AQMjQ7RCssIJynCbGUnLhQih0CJ0FT-CwwEFXxJYacYsOsrLdKzrob_rWJhsLt3nwuXZOxK53aDrRQ&passive=1209600&flowName=GlifWebSignIn&flowEntry=ServiceLogin'+','+'JASON4703',
       'https://accounts.google.com/v3/signin/identifier?dsh=S1224916245%3A1681711553453460&continue=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1IcBH4h2YkL9HBqzWFkqBIy7Qm6CTh6o%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&ec=GAZA2gE&followup=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1IcBH4h2YkL9HBqzWFkqBIy7Qm6CTh6o%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&hl=zh-TW&ifkv=AQMjQ7SIsu2PSmBMPcx3MyxQvTc4gH3HB1tCc8PNoD3TYAm5dLwRlrQdSboxJSr3RgGZRNzN-av0eA&passive=1209600&flowName=GlifWebSignIn&flowEntry=ServiceLogin'+','+'JULIAN',
       'https://accounts.google.com/v3/signin/identifier?dsh=S83959985%3A1681711596526038&continue=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D11gvbigwoSC6U_Y7z0446IiqHje9FT34%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&ec=GAZA2gE&followup=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D11gvbigwoSC6U_Y7z0446IiqHje9FT34%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&hl=zh-TW&ifkv=AQMjQ7TFQdAK4N0FyXqg0Un8xer3exnIhaVGZTSlPV3FRFVwxC_VYNqY8lEVLQzXKXCFgO8ryE4J9A&passive=1209600&flowName=GlifWebSignIn&flowEntry=ServiceLogin'+','+'SCOTT4162',
       'https://accounts.google.com/v3/signin/identifier?dsh=S-872058158%3A1681711634509817&continue=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1VFjbRpmUbUL0xGeat1d9ISw4YvJIvCg%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&ec=GAZA2gE&followup=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1VFjbRpmUbUL0xGeat1d9ISw4YvJIvCg%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&hl=zh-TW&ifkv=AQMjQ7TVZsuRtUMX90I7YNdvAxVhQdtvLBge9YW_9XORTvZgXw55JzXV8foUdrZG1oXZI0DtJqv9PA&passive=1209600&flowName=GlifWebSignIn&flowEntry=ServiceLogin'+','+'SIMONSH',
       'https://accounts.google.com/v3/signin/identifier?dsh=S-855169725%3A1681711726026711&continue=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1n1cC11Z989d_MgiYGFKsf3uYc4rmGLY%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&ec=GAZA2gE&followup=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D1n1cC11Z989d_MgiYGFKsf3uYc4rmGLY%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&hl=zh-TW&ifkv=AQMjQ7TEnaEzsAO2L0OFX-GzKMeFFPSD4KDneY5cvcnhl42DB2RvbnEDxAh8lOl58rcjfgao97th&passive=1209600&flowName=GlifWebSignIn&flowEntry=ServiceLogin'+','+'SPANELY',
       'https://accounts.google.com/v3/signin/identifier?dsh=S319449056%3A1681711664643096&continue=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D17-3hm4jUFmZ9yAwU5273uhkMzBhcVRQ%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&ec=GAZA2gE&followup=https%3A%2F%2Fwww.google.com%2Fmaps%2Fd%2Fviewer%3Fhl%3Dzh-TW%26mid%3D17-3hm4jUFmZ9yAwU5273uhkMzBhcVRQ%26ll%3D24.04740285808056%2C120.10402299999998%26z%3D7&hl=zh-TW&ifkv=AQMjQ7QvTP5wsn10jaX4xiO_Krwor5vdZe7YIBwx9GaIbXcLHt2UkoJv5hMyYYI1paOJSAWpcoeMyA&passive=1209600&flowName=GlifWebSignIn&flowEntry=ServiceLogin'+','+'WHITE5082',
       ]

#date = str(datetime.date.today()).replace('-','')
date = str(datetime.datetime.now() + datetime.timedelta(days=1)).replace('-','')[0:9]
#date = '20230815'


def do_send_email(name):
    import smtplib
    from email.mime.text import MIMEText
    from email.header import Header
    sender = ['collection@ucs.com']
    receivers = ['DI@ucs.com'] # 接收郵件
    # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
    err_msg = "[ERROR] step 4 GoogleMap update Error\nfailed name list: " +  name
    message = MIMEText(err_msg, 'plain', 'utf-8')
    message['From'] = Header('collection', 'utf-8') # 傳送者
    message['To'] = Header('DI', 'utf-8') # 接收者
    subject = '[ERROR] step 4 GoogleMap update Error'
    message['Subject'] = Header(subject, 'utf-8')
    try:
        smtpObj = smtplib.SMTP('vrh19.ucs.com')
        smtpObj.sendmail(sender, receivers, message.as_string())
        print('郵件傳送成功')
    except smtplib.SMTPException:
        print('Error: 無法傳送郵件')



    
for uri in urll:
    retry = 1
    while retry < 3:
        print('A',retry)
        try:
            print(uri.split(','))
            
            url = uri.split(',')[0]
            name = uri.split(',')[1]
            All = rf'\\fortune\UCS\DI\OC\{name}\{name}_All.xlsx'
            TSB = rf'\\fortune\UCS\DI\OC\{name}\{name}_TSB.xlsx'
           
            
            id = '10773016@gm.scu.edu.tw'
            
            password = '0000007740'
            
            driver = webdriver.Chrome(Drvfile)
            
            
            driver.get(url)
            n = driver.window_handles
            driver.switch_to.window (n[0])
            
            #輸入帳號
            i=1
            while i < 10:
                try:
                    i=i+1
                    time.sleep(3)
                    driver.find_element_by_name('identifier').send_keys(id)
                    i=10
                except:
                    pass
            
            #下一步
            driver.find_element_by_xpath('//*[@id="identifierNext"]/div/button/span').click()
            time.sleep(3)
            
            #輸入密碼
            i=1
            while i < 10:
                try:
                    i=i+1
                    time.sleep(3)
                    driver.find_element_by_name('Passwd').send_keys(password)
                    i=10
                except:
                    pass
            
            #下一步
            driver.find_element_by_xpath('//*[@id="passwordNext"]/div/button/span').click()
            
            
            
            #編輯
            i=1
            while i < 10:
                try:
                    i=i+1
                    time.sleep(3)
                    driver.find_element_by_xpath('//*[@id="legendPanel"]/div/div/div[2]/div/div/div[1]/div[4]/div/div[2]/span/span').click()
                    i=10
                except:
                    pass
            
            
            #點選圖層(UCS)選項
            #刪除這個圖層(UCS)
            i=1
            while i < 10:
                try:
                    i=i+1
                    time.sleep(5)
                    driver.find_element_by_xpath('//*[@id="ly0-layer-status"]/span/div').click()
                    driver.find_element_by_xpath('//*[@id="ly0-layer-header"]/div[3]').click()
                    driver.find_element_by_xpath("//*[text()='刪除這個圖層']").click()
                    time.sleep(2)
                    pyautogui.press('enter')
                    i=10
                except:
                    pass
            
            #點選圖層(TSB)選項
            #刪除這個圖層(TSB)
            i=1
            while i < 10:
                try:
                    i=i+1
                    time.sleep(5)
                    driver.find_element_by_xpath('//*[@id="ly1-layer-status"]/span/div').click()
                    driver.find_element_by_xpath('//*[@id="ly1-layer-header"]/div[3]').click()
                    driver.find_element_by_xpath("//*[text()='刪除這個圖層']").click()
                    time.sleep(2)
                    pyautogui.press('enter')
                    i=10
                except:
                    pass
            
            #更改名稱
            driver.refresh()
            time.sleep(20)
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(0.5)     
            
            i=1
            while i < 10:
                try:
                    i=i+1
                    time.sleep(3)
                    driver.find_element_by_xpath("//*[text()='(未命名的圖層)']").click()
                    time.sleep(2)
                    pyautogui.typewrite(name+'_'+date+'_All')
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(0.5)
                    i=10
                except:
                    pass
            
            
            
            #匯入資料
            i=1
            while i < 10:
                try:
                    i=i+1
        
                    time.sleep(3)
                    driver.find_element_by_xpath('//*[@id="ly0-layerview-import-link"]').click()
                    time.sleep(5)
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    time.sleep(2)
                    matching_iframes = soup.select('div.fFW7wc.XKSfm-Sx9Kwc-bN97Pc iframe')
                    # 遍歷找到的所有元素
                    for iframe in matching_iframes:
                        iframe_id = iframe.get('id')
                        print(iframe_id)
                    iframe = driver.find_element_by_xpath('//iframe[@id='+"'"+iframe_id+"'"+']')
                    driver.switch_to.frame(iframe)
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'span[class="VfPpkd-vQzf8d"]'))
                    )
                    element.click()
                    driver.switch_to.default_content()
                    #pyautogui.press('enter')
                    time.sleep(2)
                    pyautogui.typewrite(All)
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(0.5)
                    i=10
                except:
                    driver.refresh()
                    time.sleep(20)
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(0.5)     
                    pass
            
            driver.switch_to.default_content()
            #緯度
            #經度
            i=1
            while i < 10:
                try:
                    i=i+1
                    time.sleep(3)
                    driver.find_element_by_xpath('//*[@id="upload-checkbox-5"]/span/div').click()
                    time.sleep(0.5)
                    driver.find_element_by_xpath('//*[@id="upload-location-radio-5-0"]/div[2]/span[1]').click()
                    time.sleep(0.5)
                    
                    time.sleep(3)
                    driver.find_element_by_xpath('//*[@id="upload-checkbox-6"]/span/div').click()
                    time.sleep(0.5)
                    driver.find_element_by_xpath('//*[@id="upload-location-radio-6-1"]/div/span[1]').click()
                    time.sleep(0.5)
                    i=10
                except:
                    pass
            
            i=1
            while i < 10:
                try:
                    i=i+1
                    #繼續
                    time.sleep(3)
                    driver.find_element_by_xpath('/html/body/div[9]/div[3]/button[1]').click()
                    #優先別
                    time.sleep(3)
                    driver.find_element_by_xpath('//*[@id="upload-radio-2"]/div/span[1]').click()
                    #完成
                    time.sleep(3)
                    driver.find_element_by_xpath('/html/body/div[7]/div[3]/button[1]').click()
                    i=10
                except:
                    pass
            i=1
            while i < 10:
                try:
                    i=i+1
                    time.sleep(10)
                    #設計
                    time.sleep(3)
                    driver.find_element_by_xpath('//*[@id="ly0-layerview-stylepopup-link"]/div[2]/div').click()
                    driver.find_element_by_xpath('//*//*[@id="layer-style-popup"]/div[3]/div[1]').click()
                    #優先別
                    time.sleep(3)
                    driver.find_element_by_xpath('//*[@id="style-by-type-selector-column-str:5YSq5YWI5Yil"]/div').click()
                    #關閉畫面
                    time.sleep(3)
                    driver.find_element_by_xpath('//*[@id="layer-style-popup"]/div[1]').click()
                    time.sleep(3)
                    i=10
                except:
                    pass
            
            r = driver.find_element_by_xpath('//*[@id="ly0-layer-items-container"]').text
            print(len(r.split('\n')))
            print(r.split('\n'))
            for i in range(len(r.split('\n'))):
                s = i+i+1
                point = str(r.split('\n')[i][0:2].replace('(',''))
                print(s,point)
                time.sleep(1)
                driver.find_element_by_xpath('//*[@id="ly0-layer-items-container"]/div['+str(s)+']').click()
                time.sleep(0.5)
                driver.find_element_by_xpath('//*[@id="ly0-layer-items-container"]/div['+str(s)+']/div[3]').click()
             
                if point == '11' :
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf26"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '9' :
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf0"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '10' :
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf12"]/div').click() 
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '8' :
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf22"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '3' :
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf20"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '7' :
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf24"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '1B' :
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf2"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '6' :
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf3"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '4A' :
                    time.sleep(1)
                    driver.find_element_by_xpath("//*[text()='更多圖示']").click()
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id="map-area-container"]/div[11]/div/div[6]/button[1]').click()
                    time.sleep(2)
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    time.sleep(2)
                    matching_iframes = soup.select('div.fFW7wc.XKSfm-Sx9Kwc-bN97Pc iframe')
                    # 遍歷找到的所有元素
                    for iframe in matching_iframes:
                        iframe_id = iframe.get('id')
                        print(iframe_id)
                    iframe = driver.find_element_by_xpath('//iframe[@id='+"'"+iframe_id+"'"+']')
                    driver.switch_to.frame(iframe)
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'span[class="VfPpkd-vQzf8d"]'))
                    )
                    element.click()
                    driver.switch_to.default_content()
                    time.sleep(2)
                    pyautogui.typewrite(rf'\\fortune\UCS\DI\OC\Map_PNG\tsb_red.png')
                    time.sleep(2)
                    pyautogui.press('enter')
                    time.sleep(5)
                    driver.find_element_by_xpath("//*[text()='確定']").click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '4B' :
                    time.sleep(1)
                    driver.find_element_by_xpath("//*[text()='更多圖示']").click()
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id="map-area-container"]/div[11]/div/div[6]/button[1]').click()
                    time.sleep(2)
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    time.sleep(2)
                    matching_iframes = soup.select('div.fFW7wc.XKSfm-Sx9Kwc-bN97Pc iframe')
                    # 遍歷找到的所有元素
                    for iframe in matching_iframes:
                        iframe_id = iframe.get('id')
                        print(iframe_id)
                    iframe = driver.find_element_by_xpath('//iframe[@id='+"'"+iframe_id+"'"+']')
                    driver.switch_to.frame(iframe)
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'span[class="VfPpkd-vQzf8d"]'))
                    )
                    element.click()
                    driver.switch_to.default_content()
                    time.sleep(2)
                    pyautogui.typewrite(rf'\\fortune\UCS\DI\OC\Map_PNG\tsb_red.png')
                    time.sleep(2)
                    pyautogui.press('enter')
                    time.sleep(5)
                    driver.find_element_by_xpath("//*[text()='確定']").click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                    
                elif point == '5B' :
                    time.sleep(1)
                    driver.find_element_by_xpath("//*[text()='更多圖示']").click()
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id="map-area-container"]/div[11]/div/div[6]/button[1]').click()
                    time.sleep(2)
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    time.sleep(2)
                    matching_iframes = soup.select('div.fFW7wc.XKSfm-Sx9Kwc-bN97Pc iframe')
                    # 遍歷找到的所有元素
                    for iframe in matching_iframes:
                        iframe_id = iframe.get('id')
                        print(iframe_id)
                    iframe = driver.find_element_by_xpath('//iframe[@id='+"'"+iframe_id+"'"+']')
                    driver.switch_to.frame(iframe)
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'span[class="VfPpkd-vQzf8d"]'))
                    )
                    element.click()
                    driver.switch_to.default_content()
                    time.sleep(2)
                    pyautogui.typewrite(rf'\\fortune\UCS\DI\OC\Map_PNG\tsb_yy.png')
                    time.sleep(2)
                    pyautogui.press('enter')
                    time.sleep(5)
                    driver.find_element_by_xpath("//*[text()='確定']").click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '5A' :
                    time.sleep(1)
                    driver.find_element_by_xpath("//*[text()='更多圖示']").click()
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id="map-area-container"]/div[11]/div/div[6]/button[1]').click()
                    time.sleep(2)
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    time.sleep(2)
                    matching_iframes = soup.select('div.fFW7wc.XKSfm-Sx9Kwc-bN97Pc iframe')
                    # 遍歷找到的所有元素
                    for iframe in matching_iframes:
                        iframe_id = iframe.get('id')
                        print(iframe_id)
                    iframe = driver.find_element_by_xpath('//iframe[@id='+"'"+iframe_id+"'"+']')
                    driver.switch_to.frame(iframe)
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'span[class="VfPpkd-vQzf8d"]'))
                    )
                    element.click()
                    driver.switch_to.default_content()
                    time.sleep(2)
                    pyautogui.typewrite(rf'\\fortune\UCS\DI\OC\Map_PNG\tsb_yy.png')
                    time.sleep(2)
                    pyautogui.press('enter')
                    time.sleep(5)
                    driver.find_element_by_xpath("//*[text()='確定']").click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '1A' :
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf13"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '2' :
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf4"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '14' :
                    time.sleep(1)
                    driver.find_element_by_xpath("//*[text()='更多圖示']").click()
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id="map-area-container"]/div[11]/div/div[6]/button[1]').click()
                    time.sleep(2)
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    time.sleep(2)
                    matching_iframes = soup.select('div.fFW7wc.XKSfm-Sx9Kwc-bN97Pc iframe')
                    # 遍歷找到的所有元素
                    for iframe in matching_iframes:
                        iframe_id = iframe.get('id')
                        print(iframe_id)
                    iframe = driver.find_element_by_xpath('//iframe[@id='+"'"+iframe_id+"'"+']')
                    driver.switch_to.frame(iframe)
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'span[class="VfPpkd-vQzf8d"]'))
                    )
                    element.click()
                    driver.switch_to.default_content()
                    time.sleep(2)
                    pyautogui.typewrite(rf'\\fortune\UCS\DI\OC\Map_PNG\P1.png')
                    time.sleep(2)
                    pyautogui.press('enter')
                    time.sleep(5)
                    driver.find_element_by_xpath("//*[text()='確定']").click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '15' :
                    time.sleep(1)
                    driver.find_element_by_xpath("//*[text()='更多圖示']").click()
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id="map-area-container"]/div[11]/div/div[6]/button[1]').click()
                    time.sleep(2)
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    time.sleep(2)
                    matching_iframes = soup.select('div.fFW7wc.XKSfm-Sx9Kwc-bN97Pc iframe')
                    # 遍歷找到的所有元素
                    for iframe in matching_iframes:
                        iframe_id = iframe.get('id')
                        print(iframe_id)
                    iframe = driver.find_element_by_xpath('//iframe[@id='+"'"+iframe_id+"'"+']')
                    driver.switch_to.frame(iframe)
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'span[class="VfPpkd-vQzf8d"]'))
                    )
                    element.click()
                    driver.switch_to.default_content()
                    time.sleep(2)
                    pyautogui.typewrite(rf'\\fortune\UCS\DI\OC\Map_PNG\DBS.png')
                    time.sleep(2)
                    pyautogui.press('enter')
                    time.sleep(5)
                    driver.find_element_by_xpath("//*[text()='確定']").click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                    
                    
            #TSB   
            i=1
            while i < 10:
                try:
                    i=i+1
                    driver.refresh()
                    time.sleep(20)
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(0.5)     
                    driver.find_element_by_xpath('//*[@id="map-action-add-layer"]').click()
                    
                    i=10
                except:
                    pass
            
            #新命名
            i=1
            while i < 10:
                try:
                    i=i+1
                    time.sleep(5)
                    driver.find_element_by_xpath("//*[text()='(未命名的圖層)']").click()
                    time.sleep(2)
                    pyautogui.typewrite(name+'_'+date+'_TSB')
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(0.5)
                    i=10
                except:
                    pass
            
            if len(pd.read_excel(TSB)) == 0 :
                try:
                    time.sleep(5)
                    driver.find_element_by_xpath('//*[@id="ly1-layer-status"]/span/div').click()
                    time.sleep(3)
                except:
                    pass
                driver.close()
                retry = 3
                quit
            else:
                pass
            #匯入資料
            i=1
            while i < 10:
                try:
                    i=i+1
                    time.sleep(5)
                    driver.find_element_by_xpath('//*[@id="ly1-layerview-import-link"]').click()
                    time.sleep(5)
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    matching_iframes = soup.select('div.fFW7wc.XKSfm-Sx9Kwc-bN97Pc iframe')
                    # 遍歷找到的所有元素
                    for iframe in matching_iframes:
                        iframe_id = iframe.get('id')
                        print(iframe_id)
                    iframe = driver.find_element_by_xpath('//iframe[@id='+"'"+iframe_id+"'"+']')
                    driver.switch_to.frame(iframe)
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'span[class="VfPpkd-vQzf8d"]'))
                    )
                    element.click()
                    #pyautogui.press('enter')
                    time.sleep(2)
                    pyautogui.typewrite(TSB)
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(0.5)     
                    i=10
                except:
                    driver.refresh()
                    time.sleep(20)
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(0.5)     
                    pass
            driver.switch_to.default_content()
            i=1
            while i < 10:
                try:
                    i=i+1
                    #緯度
                    time.sleep(10)
                    driver.find_element_by_xpath('//*[@id="upload-checkbox-5"]/span/div').click()
                    time.sleep(0.5)
                    driver.find_element_by_xpath('//*[@id="upload-location-radio-5-0"]/div[2]/span[1]').click()
                    time.sleep(0.5)
                    #經度
                    time.sleep(3)
                    driver.find_element_by_xpath('//*[@id="upload-checkbox-6"]/span/div').click()
                    time.sleep(0.5)
                    driver.find_element_by_xpath('//*[@id="upload-location-radio-6-1"]/div/span[1]').click()
                    time.sleep(0.5)
                    #繼續
                    driver.find_element_by_xpath('/html/body/div[9]/div[3]/button[1]').click()     
                    #countdown
                    driver.find_element_by_xpath('//*[@id="upload-radio-0"]/div/span[1]').click()
                    #完成
                    driver.find_element_by_xpath('/html/body/div[7]/div[3]/button[1]').click()
                    
                    i=10
                except:
                    pass 
                    
                   
            i=1
            while i < 10:
                try:
                    i=i+1
                    time.sleep(10)
                    #設計
                    driver.find_element_by_xpath('//*[@id="ly1-layerview-stylepopup-link"]/div[2]/div').click()
                    driver.find_element_by_xpath('//*//*[@id="layer-style-popup"]/div[3]/div[1]').click()
                    
                    #countdown
                    try:
                        driver.find_element_by_xpath('//*[@id="style-by-type-selector-column-double:Y291bnRkb3du"]/div').click()
                    except:
                        driver.find_element_by_xpath('//*[@id="style-by-type-selector-column-str:Y291bnRkb3du"]/div').click()
                    
                    
                    #關閉畫面
                    driver.find_element_by_xpath('//*[@id="layer-style-popup"]/div[1]').click()     
                    i=10
                except:
                    pass 
            
            
            r = driver.find_element_by_xpath('//*[@id="ly1-layer-items-container"]').text
            
            
            print(len(r.split('\n')))
            print(r.split('\n'))
            for i in range(len(r.split('\n'))):
                s = i+i+1
                point = str(r.split('\n')[i][0:2].replace('(',''))
                print(s,point)
                time.sleep(1)
                driver.find_element_by_xpath('//*[@id="ly1-layer-items-container"]/div['+str(s)+']').click()
                driver.find_element_by_xpath('//*[@id="ly1-layer-items-container"]/div['+str(s)+']/div[3]').click()
                if '其他' in point :
                    time.sleep(1)
                    driver.find_element_by_xpath("//*[text()='更多圖示']").click()
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id="map-area-container"]/div[11]/div/div[6]/button[1]').click()
                    time.sleep(2)
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    time.sleep(2)
                    matching_iframes = soup.select('div.fFW7wc.XKSfm-Sx9Kwc-bN97Pc iframe')
                    # 遍歷找到的所有元素
                    for iframe in matching_iframes:
                        iframe_id = iframe.get('id')
                        print(iframe_id)
                    iframe = driver.find_element_by_xpath('//iframe[@id='+"'"+iframe_id+"'"+']')
                    driver.switch_to.frame(iframe)
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'span[class="VfPpkd-vQzf8d"]'))
                    )
                    element.click()
                    driver.switch_to.default_content()
                    time.sleep(2)
                    pyautogui.typewrite(rf'\\fortune\UCS\DI\OC\Map_PNG\tsb_blue.psd2.png')
                    time.sleep(2)
                    pyautogui.press('enter')
                    time.sleep(5)
                    driver.find_element_by_xpath("//*[text()='確定']").click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif '尚未' in point :
                    time.sleep(1)
                    driver.find_element_by_xpath("//*[text()='更多圖示']").click()
                    time.sleep(1)
                    driver.find_element_by_xpath('//*[@id="map-area-container"]/div[11]/div/div[6]/button[1]').click()
                    time.sleep(2)
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    time.sleep(2)
                    matching_iframes = soup.select('div.fFW7wc.XKSfm-Sx9Kwc-bN97Pc iframe')
                    # 遍歷找到的所有元素
                    for iframe in matching_iframes:
                        iframe_id = iframe.get('id')
                        print(iframe_id)
                    iframe = driver.find_element_by_xpath('//iframe[@id='+"'"+iframe_id+"'"+']')
                    driver.switch_to.frame(iframe)
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'span[class="VfPpkd-vQzf8d"]'))
                    )
                    element.click()
                    driver.switch_to.default_content()
                    time.sleep(2)
                    pyautogui.typewrite(rf'\\fortune\UCS\DI\OC\Map_PNG\tsb_blue.psd2.png')
                    time.sleep(2)
                    pyautogui.press('enter')
                    time.sleep(5)
                    driver.find_element_by_xpath("//*[text()='確定']").click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '0' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf1"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '1' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf2"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '2' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf3"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '3' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf4"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '4' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf5"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '5' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf6"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '6' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf7"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '7' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf8"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '8' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf9"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '9' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf10"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '10' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf11"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '11' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf12"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '12' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf13"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '13' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf14"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '14' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf15"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '15' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf16"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '16' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf17"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '17' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf18"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '18' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf19"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '19' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf20"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                elif point == '20' :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf21"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                else :
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id=":11"]/div/div/div').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="VIpgJd-nEeMgc-eEDwDf29"]/div').click()
                    driver.find_element_by_xpath('//*[@id="stylepopup-close"]').click()
                    
            time.sleep(2)
            driver.find_element_by_xpath('//*[@id="ly1-layer-status"]/span/div').click()
            time.sleep(2)
            driver.close()
            retry = 3
            print('B',retry)
        except:
            retry = retry +1
            os.system('pkill chrome')
            os.system("kill $(ps aux | grep webdriver| awk '{print $2}')")
            os.system('taskkill /F /IM chrome.exe')
            print('C',retry)
            if retry == 3 :
                os.system('pkill chrome')
                os.system("kill $(ps aux | grep webdriver| awk '{print $2}')")
                os.system('taskkill /F /IM chrome.exe')
                do_send_email(name)
                quit
            else:
                pass
    















        
        
        