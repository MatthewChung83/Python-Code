# -*- coding: utf-8 -*-
"""
Created on Thu Feb  2 11:49:29 2023

@author: admin
"""
db = {
    'server': 'vnt07.ucs.com',
    'database': 'UIS',
    'username': 'pyuser',
    'password': 'Ucredit7607',
}

server,database,username,password= db['server'],db['database'],db['username'],db['password']

def table(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select *
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] 
        where data_date = convert(varchar(10),getdate(),111)
        
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src
#a = table(server,username,password,database)

import time
i = 1
while i < 10:
    a = table(server,username,password,database)
    if len(a) ==0:
        time.sleep(300)
        i = i+1
        print(i)
        #Send Mail        
        import smtplib
        from email.mime.text import MIMEText
        from email.header import Header
        sender = ['collection@ucs.com']
        receivers = ['DI@ucs.com'] # 接收郵件
        # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
        message = MIMEText('[ERROR] step 1 SVTb_PrioritySV_LIST 補經緯度 skiptrace.dbo.SVTb_PrioritySV_LIST資料異常', 'plain', 'utf-8')
        message['From'] = Header('collection', 'utf-8') # 傳送者
        message['To'] = Header('DI', 'utf-8') # 接收者
        subject = '[ERROR] step 1 SVTb_PrioritySV_LIST 補經緯度 '
        message['Subject'] = Header(subject, 'utf-8')
        try:
            smtpObj = smtplib.SMTP('vrh19.ucs.com')
            smtpObj.sendmail(sender, receivers, message.as_string())
            print('郵件傳送成功')
        except smtplib.SMTPException:
            print('Error: 無法傳送郵件')
    else:
        i=10
        pass