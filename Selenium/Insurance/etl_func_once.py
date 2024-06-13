# -*- coding: utf-8 -*-
"""
Created on Mon May  9 18:38:29 2022

@author: admin
"""


def num(obs,server,username,password,database,fromtb,totb,check):
    obs = src_obs(server,username,password,database,fromtb,totb,check)
    if obs > 0 :
        import pymssql
        conn = pymssql.connect(server=server, user=username, password=password, database = database)
        cursor = conn.cursor()
        
        script = f"""
        update UCS_Source
        set F_OBS = '{obs}'
        where Engine = 'insurance' and type = '{check}'
        """
        cursor.execute(script)
        conn.commit()
        cursor.close()
        conn.close()
    elif obs ==0 :
        import pymssql
        conn = pymssql.connect(server=server, user=username, password=password, database = database)
        cursor = conn.cursor()
        
        script = f"""
        update UCS_Source
        set status = 'Done'
        where Engine = 'insurance' and type = '{check}'
        """
        cursor.execute(script)
        conn.commit()
        cursor.close()
        conn.close()
        
def check(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select  type
    from UCS_Source  where  Engine='insurance' and  status <> 'Done'  
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return list(c_src[0])[0]

def CAPTCHAOCR(Apurl,imgf,imgp):
    import requests
    url = Apurl
    
    files = {
        "image_file":(imgf,open(imgp,"rb"),'image/jpeg')
    }
    response = requests.request("POST", url, files=files)
    return response.text
def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num
def src_obs(server,username,password,database,fromtb,totb,check):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""

    select count(*) from [{fromtb}]  b
	left join [{totb}] r on b.ID = r.ID 
	where r.id in (select pid from treasure.skiptrace.dbo.WebScraping where WebItem like '%保險業務員查詢%' and flag = 'Y' and note = '未辦理登錄') and update_date < getdate()-1
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]
def insurance(doc):
    insurance= []

    insurance.append({
        
        "Name":doc[0],
        "ID":doc[1],
        "birthday":doc[2],
        "insurance_num":doc[3],
        "update_date":doc[4],
        "note":doc[5],
        
    })
    return insurance
def toSQL(docs, totb, server, database, username, password):
    import pyodbc
    #conn_cmd = 'DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password
    conn_cmd = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password
    with pyodbc.connect(conn_cmd) as cnxn:
        cnxn.autocommit = False
        with cnxn.cursor() as cursor:
            data_keys = ','.join(docs[0].keys())
            data_symbols = ','.join(['?' for _ in range(len(docs[0].keys()))])
            insert_cmd = """INSERT INTO {} ({})
            VALUES ({})""".format(totb,data_keys,data_symbols )
            data_values = [tuple(doc.values()) for doc in docs]
            cursor.executemany( insert_cmd,data_values )
            cnxn.commit()


def dbfrom(server,username,password,database,fromtb,totb,check):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""

    select b.* ,r.rowid
    into #test
    from {fromtb} b
	left join [{totb}] r on b.ID = r.ID 
	where r.id in (select pid from treasure.skiptrace.dbo.WebScraping where WebItem like '%保險業務員查詢%' and flag = 'Y' and note = '未辦理登錄') and update_date < getdate()-1
    select * from #test order by personi
    --order by personi
	offset 0 row fetch next 1 rows only
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def mail(obs):
    if obs ==0:
        import smtplib
        from email.mime.text import MIMEText
        from email.header import Header
        sender = 'matthew5043@ucs.com'
        receivers = ['matthew5043@ucs.com'] # 接收郵件
        # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
        message = MIMEText('insurance end', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'insurance end'
        message['Subject'] = Header(subject, 'utf-8')
        try:
            smtpObj = smtplib.SMTP('vrh19.ucs.com')
            smtpObj.sendmail(sender, receivers, message.as_string())
            print('郵件傳送成功')
        except smtplib.SMTPException:
            print('Error: 無法傳送郵件')
    else:
        import smtplib
        from email.mime.text import MIMEText
        from email.header import Header
        sender = 'matthew5043@ucs.com'
        receivers = ['matthew5043@ucs.com'] # 接收郵件
        # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
        message = MIMEText('insurance start', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'insurance start'
        message['Subject'] = Header(subject, 'utf-8')
        try:
            smtpObj = smtplib.SMTP('vrh19.ucs.com')
            smtpObj.sendmail(sender, receivers, message.as_string())
            print('郵件傳送成功')
        except smtplib.SMTPException:
            print('Error: 無法傳送郵件')
def errormail(obs):
    if obs ==0:
        import smtplib
        from email.mime.text import MIMEText
        from email.header import Header
        sender = 'matthew5043@ucs.com'
        receivers = ['matthew5043@ucs.com'] # 接收郵件
        # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
        message = MIMEText('insurance end', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'insurance end'
        message['Subject'] = Header(subject, 'utf-8')
        try:
            smtpObj = smtplib.SMTP('vrh19.ucs.com')
            smtpObj.sendmail(sender, receivers, message.as_string())
            print('郵件傳送成功')
        except smtplib.SMTPException:
            print('Error: 無法傳送郵件')
    else:
        import smtplib
        from email.mime.text import MIMEText
        from email.header import Header
        sender = 'matthew5043@ucs.com'
        receivers = ['matthew5043@ucs.com'] # 接收郵件
        # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
        message = MIMEText('insurance ERROR', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'insurance ERROR'
        message['Subject'] = Header(subject, 'utf-8')
        try:
            smtpObj = smtplib.SMTP('vrh19.ucs.com')
            smtpObj.sendmail(sender, receivers, message.as_string())
            print('郵件傳送成功')
        except smtplib.SMTPException:
            print('Error: 無法傳送郵件')
def CAPTCHAOCR(Apurl,imgf,imgp):
    import requests
    url = Apurl
    
    files = {
        "image_file":(imgf,open(imgp,"rb"),'image/jpeg')
    }
    response = requests.request("POST", url, files=files)
    return response.text

def update(server,username,password,database,totb,note,ID,today,rowid,insurance_num):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update [{totb}]
    set note = '{note}',update_date = '{today}',insurance_num = '{insurance_num}'
    where id = '{ID}' and rowid = '{rowid}'
    
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()