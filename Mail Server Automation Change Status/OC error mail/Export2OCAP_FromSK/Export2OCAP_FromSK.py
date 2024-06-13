# -*- coding: utf-8 -*-
"""
Created on Tue Jun  7 09:24:30 2022

@author: matthew5043
"""

#Send Mail        
import smtplib
from email.mime.text import MIMEText
from email.header import Header
sender = 'collection@ucs.com'
receivers = ['matthew5043@ucs.com'] # 接收郵件
# 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
message = MIMEText('Job-03- 每日 (23:00) 匯入外訪平台 Step 1 外訪平台資料自Treasure.skiptrace匯入至10.90.0.194.OCAP exec Export2OCAP_FromSK', 'plain', 'utf-8')
message['From'] = Header('collection', 'utf-8') # 傳送者
message['To'] = Header('SW', 'utf-8') # 接收者
subject = 'VNT07 ERROR Job-03- 每日 (23:00) 匯入外訪平台 Step 1 外訪平台資料自Treasure.skiptrace匯入至10.90.0.194.OCAP'
message['Subject'] = Header(subject, 'utf-8')
try:
    smtpObj = smtplib.SMTP('vrh19.ucs.com')
    smtpObj.sendmail(sender, receivers, message.as_string())
    print('郵件傳送成功')
except smtplib.SMTPException:
    print('Error: 無法傳送郵件')