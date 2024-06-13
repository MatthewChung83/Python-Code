import io
from symbol import except_clause
import time 
import calendar
import requests
#import pyautogui
import os

from PIL import Image
from bs4 import BeautifulSoup
from selenium import webdriver
import selenium.common.exceptions

from PIL import ImageGrab
import pandas as pd
import datetime

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.expected_conditions import alert_is_present

import imaplib
import email
from email.header import decode_header
import re


from config import *
from etl_func import *
from dict import *

server,database,username,password,totb,fromtb = db['server'],db['database'],db['username'],db['password'],db['totb'],db['fromtb']
url,Drvfile= wbinfo['url'],wbinfo['Drvfile']
Name001,Mail001,PSD001 = vars['Name001'],vars['Mail001'],vars['PSD001']
Name002,Mail002,PSD002 = vars['Name002'],vars['Mail002'],vars['PSD002']
Name003,Mail003,PSD003 = vars['Name003'],vars['Mail003'],vars['PSD003']
Phone,Compiled,Company,Address = vars['Phone'],vars['Compiled'],vars['Company'],vars['Address']
#chrome_options = Options()
#chrome_options.add_argument('--headless')
#driver = webdriver.Chrome(Drvfile,chrome_options=chrome_options)
today = datetime.datetime.now().strftime('%Y-%m-%d')
obs = src_obs(server,username,password,database,fromtb,totb)
file_date = datetime.datetime.now().strftime('%Y%m%d')
try:
    os.makedirs(rf'C:\Py_Project\project\Legal_insur\file\image\{str(file_date)}')
    #os.makedirs(rf'\\vnt07\ucsfile\el\{entitytype}')
    #os.makedirs(rf'//fortune/Cashfile/UCS/EL/{weekly}/')
except:
    pass
while obs > 0 :
    try:
        for i in foo(-1,obs-1):
            driver = webdriver.Chrome(Drvfile)
            driver.get(url)
            driver.maximize_window()
            n = driver.window_handles
            driver.switch_to.window (n[0])  
            src = dbfrom(server,username,password,database,fromtb,totb,today)[0]
            UUID = str(src[0])
            Casei = str(src[3])
            legal_num = str(src[5])
            legal_court = str(src[6])
            ID = str(src[7])
            Insur_Type = str(src[17])
            RequestDate = str(src[21])
            print(RequestDate)
            print(today)
            if RequestDate == today :
                print('A')
                pass
            else :
                quit()
            if Insur_Type == '1' :
                Name = str(Name001)
                Mail = str(Mail001)
                PSD = str(PSD001)
            elif Insur_Type == '2' :
                Name = str(Name002)
                Mail = str(Mail002)
                PSD = str(PSD002)
            elif Insur_Type == '3' :
                Name = str(Name003)
                Mail = str(Mail003)
                PSD = str(PSD003)
            else:
                quit()
            #輸入聯絡人姓名
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME,'contactName')))
            driver.find_element_by_xpath('//*[@id="submitForm"]/ul/li[1]/input').send_keys(Name)
            #輸入聯絡人電話
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME,'contactPersonPhone')))
            driver.find_element_by_xpath('//*[@id="submitForm"]/ul/li[2]/input').send_keys(Phone)
            #輸入聯絡人信箱
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME,'contactEmail')))
            driver.find_element_by_xpath('//*[@id="contactEmail"]').send_keys(Mail)
            #點選寄送驗證碼
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID,'otpVerificationCodeBtn')))
            driver.find_element_by_xpath('//*[@id="otpVerificationCodeBtn"]').click()
            time.sleep(3)
            #點選OK
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID,'popupTitle')))
            time.sleep(3)
            driver.find_element_by_xpath('//*[@id="popupbox"]/div/div/div[3]/a').click()
            time.sleep(3)
            captcha = ''
            while len(captcha) == 0:
                #找尋最新驗證碼
                # 登錄信息
                #username = 'ucsin001@gmail.com'
                #password = 'hrub xgea gnko ojwl'
                # 建立與 Gmail 的連接
                mail = imaplib.IMAP4_SSL("imap.gmail.com")
                mail.login(Mail, PSD)
                # 選擇郵箱，'INBOX' 表示收件箱
                mail.select('inbox')
                # 搜索郵件
                status, messages = mail.search(None,'ALL')
                # 如果要讀取所有郵件，可以將條件改為 'ALL'
                if status != 'OK':
                    print("沒有找到郵件！")
                else:
                    # messages 是一個郵件編號列表，這裡只取最新的一封
                    latest_email_id = messages[0].split()[-1]
                    status, data = mail.fetch(latest_email_id, '(RFC822)')

                    if status != 'OK':
                        print("無法讀取郵件！")
                    else:
                        # 解析郵件內容
                        msg = email.message_from_bytes(data[0][1])
                        #print(msg)
                        if msg.is_multipart():
                            for part in msg.walk():
                                if part.get_content_type() == "text/plain" or part.get_content_type() == "text/html":
                                    message = part.get_payload(decode=True).decode()
                                    #print(message)
                                    # 使用正則表達式尋找驗證碼
                                    match = re.search(r'驗證碼為 (\d{6})', message)
                                    if match:
                                        print('找到的驗證碼:', match.group(1))
                                        break
                        else:
                            # 非多部分郵件，直接搜尋驗證碼
                            message = msg.get_payload(decode=True).decode()
                            #print(message)
                            match = re.search(r'驗證碼為 <span>(\d{6})</span>', message)
                            #print(match)
                            if match:
                                print('找到的驗證碼:', match.group(1))
                                captcha = match.group(1)
                # 斷開連接
                mail.logout()
                

            #輸入信箱驗證碼
            driver.find_element_by_xpath('//*[@id="verifyCodeInput"]').send_keys(captcha)
            #輸入公務機關案號
            driver.find_element_by_xpath('//*[@id="submitForm"]/ul/li[5]/input').send_keys(legal_num)
            #輸入公文來源
            driver.find_element_by_xpath('//*[@id="submitForm"]/ul/li[6]/input').send_keys(legal_court)
            #輸入被查詢人身分證
            for i in range(len(ID.split(','))):

                if i == 0:
                    driver.find_element_by_xpath('//*[@id="queryId1"]').send_keys(ID.split(',')[i])
                else:
                    driver.find_element_by_xpath('//*[@id="submitForm"]/ul/button').click()
                    driver.find_element_by_xpath('//*[@id="appendQueryIdInputList"]/li['+str(int(i)+1)+']/input').send_keys(ID.split(',')[i])
            time.sleep(3)       
            #本人確認即送出
            driver.find_element_by_xpath('//*[@id="submitForm"]/ul/li[7]/div/label/span').click()
            time.sleep(3)
            driver.find_element_by_xpath('//*[@id="submitBtn"]/a').click()
            time.sleep(3)
            try:
                Notes2 = driver.find_element_by_id('popupMessage2').text
            except:
                Notes2 = ''
            try:
                Notes = driver.find_element_by_id('popupMessage').text
                if '查無資料' in Notes :#and len(Notes2) == 0:
                    Order_Num = ''
                    Account_Name = ''
                    Payment_Type = ''
                    Product_List = ''
                    Transfer_Bank = ''
                    Transfer_Account = ''
                    Payment_Deadline = ''
                    Transfer_Fee = ''
                    Legal_Type = '保險查詢費'
                    Status = 'F'
                    update(server,username,password,database,totb,Notes,Casei,Order_Num,Account_Name,Payment_Type,Product_List,Transfer_Bank,Transfer_Account
                    ,Payment_Deadline,Transfer_Fee,Legal_Type,Status,today)
                    time.sleep(3)
                    #點選OK
                    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID,'popupTitle')))
                    driver.find_element_by_xpath('//*[@id="popupbox"]/div/div/div[3]/a').click()
                    time.sleep(3)
                    driver.refresh()
                elif '該案' in Notes:
                    Order_Num = ''
                    Account_Name = ''
                    Payment_Type = ''
                    Product_List = ''
                    Transfer_Bank = ''
                    Transfer_Account = ''
                    Payment_Deadline = ''
                    Transfer_Fee = ''
                    Legal_Type = '保險查詢費'
                    Status = 'F'
                    update(server,username,password,database,totb,Notes,Casei,Order_Num,Account_Name,Payment_Type,Product_List,Transfer_Bank,Transfer_Account
                    ,Payment_Deadline,Transfer_Fee,Legal_Type,Status,today)
                    time.sleep(3)
                    #點選OK
                    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID,'popupTitle')))
                    driver.find_element_by_xpath('//*[@id="popupbox"]/div/div/div[3]/a').click()
                    time.sleep(3)
                    driver.refresh()
                #elif '查無資料' in Notes and len(Notes2) == 0 :
                #    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID,'popupTitle')))
                #    driver.find_element_by_xpath('//*[@id="popupbox"]/div/div/div[3]/a').click()
                #    time.sleep(3)
                #    driver.refresh()
                #    continue
                else:
                    driver.find_element_by_xpath('//*[@id="popupbox"]/div/div/div[3]/a').click()
                    driver.refresh()
                    os.system('pkill chrome')
                    os.system("kill $(ps aux | grep webdriver| awk '{print $2}')")
                    os.system('taskkill /F /IM chrome.exe')
                    time.sleep(60)
                    continue
            except:
                Notes = '調閱成功'
                #如果有資料
                #統一編號
                WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID,'taxIdNumber')))
                driver.find_element_by_xpath('//*[@id="taxIdNumber"]').send_keys(Compiled)
                #公司名稱
                driver.find_element_by_xpath('//*[@id="companyName"]').send_keys(Company)
                #公司地址
                driver.find_element_by_xpath('//*[@id="companyAddress"]').send_keys(Address)
                #ATM繳費
                driver.find_element_by_xpath('//*[@id="pageSize"]/div[2]/div/div[2]/ul/li[3]/a').click()
                #點選中信
                WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID,'selATMBank')))
                Select(driver.find_element_by_id('selATMBank')).select_by_visible_text(u'中國信託')
                time.sleep(3)
                #取得繳費帳號
                driver.find_element_by_xpath('//*[@id="ATMPaySubmit"]').click()
                time.sleep(5)
                #訂單編號
                Order_Num = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[2]/dl[1]/dd').text
                #商店名稱
                Account_Name = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[2]/dl[2]/dd').text
                #付款方式
                Payment_Type = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[2]/dl[3]/dd').text
                #商品明細
                Product_List = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[3]/div[1]/dl[2]/dd[1]').text
                #繳費帳號#銀行代碼
                Transfer_Bank = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[4]/dl[1]/dd/p[1]').text
                Transfer_Account = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[4]/dl[1]/dd/p[2]/span').text
                Transfer_Account = Transfer_Account.replace(' ','')
                #繳費截止時間
                Payment_Deadline = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[4]/dl[2]/dd').text
                #缺少繳費金額
                Transfer_Fee = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[3]/div[1]/dl[2]/dd[2]').text
                #法務類別
                Legal_Type = '保險查詢費'
                #狀態
                Status = 'F'
                
                #更新系統
                print(server,username,password,database,totb,Notes,UUID,Order_Num,Account_Name,Payment_Type,Product_List,Transfer_Bank,Transfer_Account
                    ,Payment_Deadline,Transfer_Fee,Legal_Type,Status)
                update(server,username,password,database,totb,Notes,Casei,Order_Num,Account_Name,Payment_Type,Product_List,Transfer_Bank,Transfer_Account
                    ,Payment_Deadline,Transfer_Fee,Legal_Type,Status,today)
                time.sleep (2)
                #申請截圖
                im = ImageGrab.grab()
                PEL1 = rf'{Casei}.jpg'
                im.save(rf'C:\Py_Project\project\Legal_insur\file\image\{str(file_date)}\{str(PEL1)}')
                time.sleep (2)
            os.system('pkill chrome')
            os.system("kill $(ps aux | grep webdriver| awk '{print $2}')")
            os.system('taskkill /F /IM chrome.exe')
        obs = src_obs(server,username,password,database,fromtb,totb)
    except:
        obs = src_obs(server,username,password,database,fromtb,totb)