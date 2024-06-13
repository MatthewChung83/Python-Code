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

wbinfo = {'Drvfile':rf'D:\chromedriver-win32\chromedriver.exe'}

Drvfile = wbinfo['Drvfile']

url = 'https://insurtech.lia-roc.org.tw/crr/index.html'

driver = webdriver.Chrome(Drvfile)
driver.get(url)

Name = '錢先生'
Phone = '09123456789'
Mail = 'ucsin001@gmail.com'
captcha = ''
legal_num = '18745'
legal_court = '臺灣士林地方法院民事執行處'
ID = 'F123443170'
Compiled = '23756020'
Company = '聯合財信資產管理股份有限公司'
Address = '台北市北投區裕民六路2號3樓'

#輸入聯絡人姓名
driver.find_element_by_xpath('//*[@id="submitForm"]/ul/li[1]/input').send_keys(Name)
#輸入聯絡人電話
driver.find_element_by_xpath('//*[@id="submitForm"]/ul/li[2]/input').send_keys(Phone)
#輸入聯絡人信箱
driver.find_element_by_xpath('//*[@id="contactEmail"]').send_keys(Mail)
#點選寄送驗證碼
driver.find_element_by_xpath('//*[@id="otpVerificationCodeBtn"]').click()
#點選OK
driver.find_element_by_xpath('//*[@id="popupbox"]/div/div/div[3]/a').click()

#找尋最新驗證碼
import imaplib
import email
from email.header import decode_header
import re

# 登錄信息
username = 'ucsin001@gmail.com'
password = 'hrub xgea gnko ojwl'

# 建立與 Gmail 的連接
mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login(username, password)

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
driver.find_element_by_xpath('//*[@id="queryId1"]').send_keys(ID)
#本人確認即送出
driver.find_element_by_xpath('//*[@id="submitForm"]/ul/li[7]/div/label/span').click()
driver.find_element_by_xpath('//*[@id="submitBtn"]/a').click()
status = driver.find_element_by_id('popupMessage').text
#如果有資料
#統一編號
driver.find_element_by_xpath('//*[@id="taxIdNumber"]').send_keys(Compiled)
#公司名稱
driver.find_element_by_xpath('//*[@id="companyName"]').send_keys(Company)
#公司地址
driver.find_element_by_xpath('//*[@id="companyAddress"]').send_keys(Address)
#ATM繳費
driver.find_element_by_xpath('//*[@id="pageSize"]/div[2]/div/div[2]/ul/li[3]/a').click()
#點選中信
driver.find_element_by_xpath('//*[@id="selATMBank"]/option[6]').click()
#取得繳費帳號
driver.find_element_by_xpath('//*[@id="ATMPaySubmit"]').click()
#訂單編號
print(driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[2]/dl[1]/dd').text)
#商店名稱
print(driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[2]/dl[2]/dd').text)
#付款方式
print(driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[2]/dl[3]/dd').text)
#商品明細
print(driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[3]/div[1]/dl[2]/dd[1]').text)
#繳費帳號#銀行代碼
print(driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[4]/dl[1]/dd/p[1]').text)
print(driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[4]/dl[1]/dd/p[2]/span').text)
#繳費截止時間
print(driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[4]/dl[2]/dd').text)