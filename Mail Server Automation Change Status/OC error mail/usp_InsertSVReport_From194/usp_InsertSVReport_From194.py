# -*- coding: utf-8 -*-
"""
Created on Tue Jun  7 10:01:03 2022

@author: matthew5043
"""

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
message = MIMEText('Job-02- 每日 (06:00) 資訊匯入 SK Step 1 外訪報告匯入 exec usp_InsertSVReport_From194', 'plain', 'utf-8')
message['From'] = Header('collection', 'utf-8') # 傳送者
message['To'] = Header('SW', 'utf-8') # 接收者
subject = 'VNT07 ERROR Job-02- 每日 (06:00) 資訊匯入 SK Step 1 外訪報告匯入'
message['Subject'] = Header(subject, 'utf-8')
try:
    smtpObj = smtplib.SMTP('vrh19.ucs.com')
    smtpObj.sendmail(sender, receivers, message.as_string())
    print('郵件傳送成功')
except smtplib.SMTPException:
    print('Error: 無法傳送郵件')