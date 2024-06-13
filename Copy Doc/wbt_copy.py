# -*- coding: utf-8 -*-
"""
Created on Wed Nov 30 16:07:36 2022

@author: admin
"""
from pathlib import Path
from bs4 import BeautifulSoup
from urllib.parse import urlencode
import requests

db = {
    'server': '10.10.0.94',
    'database': 'CL_Daily',
    'username': 'CLUSER',
    'password': 'Ucredit7607',
    #'fromtb':'base_case',
    #'totb':'resim',
    #'checktb':'UCS_Source',
    
}
server,database,username,password = db['server'],db['database'],db['username'],db['password']
import datetime

def getYesterday(): 
    yesterday = datetime.date.today() + datetime.timedelta(-6)
    return yesterday
yesterday = getYesterday()
print(yesterday)


def src_obs(server,username,password,database,yesterday):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*) from wbt_court_auction_tb
    where entrydate > = '{yesterday}' 
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]

def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num
def dbfrom(server,username,password,database,yesterday):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select court+'_'+number+'_'+replace(replace(dbo.SplitX(date, ' ',0),'/',''),' ','')+'_'+turn+'.pdf',document from wbt_court_auction_tb
    where entrydate > = '{yesterday}' order by entrydate
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

src = dbfrom(server,username,password,database,yesterday)
src_obs = src_obs(server,username,password,database,yesterday)
#print(src_obs)
#print(src)


print(len(src))
import pyautogui
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.expected_conditions import alert_is_present
from selenium import webdriver
import selenium.common.exceptions
import pdfkit
import time
import pyperclip
from pathlib import Path
from bs4 import BeautifulSoup
from urllib.parse import urlencode
import requests
Drvfile = rf'C:\Py_Project\env\chromedriver_win32\chromedriver'
for i in src:
    url = i[1]
    print(url)
    file = i[0]
    print(file)
    ta_file = rf'\\fortune\Cashfile\UCS\WBT\\'+str(file)

    
    #html = url 
    #config = pdfkit.configuration(wkhtmltopdf="C:\Program Files\wkhtmltopdf\wkhtmltopdf.exe")
    #pdfkit_options = {'encoding':"big5"}
    #file_save =  ta_file
    #pdfkit.from_url(html,file_save,configuration=config,options= pdfkit_options)
    try:
        filename = ta_file
        filename1 = Path(filename)
        url3 = url
        response = requests.get(url3)
        filename1.write_bytes(response.content)
    except:
        pass
#import shutil
#import os
#import time
#import datetime


#file_source = rf'C:\Py_Project\tfasc\file\\'
#file_destination = rf'\\fortune\Cashfile\UCS\WBT\\'
 
#get_files = os.listdir(file_source)

#time.strftime('%Y%m%d')
#now_time = datetime.datetime.now()
#future_time = now_time + datetime.timedelta(days=-4)
#fu = future_time.strftime('%Y%m%d')
#print(fu)

#for g in get_files:
#    t = time.localtime(os.path.getctime(file_source+g))
#    ctime=time.strftime("%Y%m%d",t)
    
#    if ctime > fu :     
#        try:
#            shutil.copy(file_source + g, file_destination)
#            print(ctime)
#            print(g)
#            pass
#        except:
#            pass
    
#    else :
#        pass
#Send Mail        
import smtplib
from email.mime.text import MIMEText
from email.header import Header
sender = 'matthew5043@ucs.com'
receivers = ['matthew5043@ucs.com'] # 接收郵件
# 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
message = MIMEText('法拍上載圖檔已完成', 'plain', 'utf-8')
message['From'] = Header('wbt', 'utf-8') # 傳送者
message['To'] = Header('matthew5043', 'utf-8') # 接收者
subject = '法拍上載圖檔已完成'
message['Subject'] = Header(subject, 'utf-8')
try:
    smtpObj = smtplib.SMTP('vrh19.ucs.com')
    smtpObj.sendmail(sender, receivers, message.as_string())
    print('郵件傳送成功')
except smtplib.SMTPException:
    print('Error: 無法傳送郵件')