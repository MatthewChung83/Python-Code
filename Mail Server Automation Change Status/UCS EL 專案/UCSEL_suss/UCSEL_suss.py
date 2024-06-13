# -*- coding: utf-8 -*-
"""
Created on Wed Jun 15 17:04:14 2022

@author: matthew5043
"""

import smtplib
from email.mime.text import MIMEText
from email.header import Header
sender = 'collection@ucs.com'
receivers = ['matthew5043@ucs.com'] # 接收郵件
# 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
message = MIMEText('UCS EL 專案已完成', 'plain', 'utf-8')
message['From'] = Header('UCSEL', 'utf-8') # 傳送者
message['To'] = Header('matthew', 'utf-8') # 接收者
subject = 'UCS EL 專案已完成'
message['Subject'] = Header(subject, 'utf-8')
try:
    smtpObj = smtplib.SMTP('vrh19.ucs.com')
    smtpObj.sendmail(sender, receivers, message.as_string())
    print('郵件傳送成功')
except smtplib.SMTPException:
    print('Error: 無法傳送郵件')