# -*- coding: utf-8 -*-
"""
Created on Wed Mar  2 16:46:20 2022

@author: admin
"""




import shutil
import os
import time
import datetime


file_source = rf'C:\Py_Project\project\court-auction-crawler-0.1.3\court-auction-crawler-0.1.3\crawler\data\\'
file_destination = rf'\\fortune\Cashfile\UCS\WBT\\'
 
get_files = os.listdir(file_source)

time.strftime('%Y%m%d')
now_time = datetime.datetime.now()
future_time = now_time + datetime.timedelta(days=-6)
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