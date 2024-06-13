# -*- coding: utf-8 -*-
"""
Created on Mon May  9 18:34:52 2022

@author: admin
"""


def mail(obs):
    if obs ==0:
        import smtplib
        from email.mime.text import MIMEText
        from email.header import Header
        sender = 'matthew5043@ucs.com'
        receivers = ['matthew5043@ucs.com'] # 接收郵件
        # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
        message = MIMEText('judicial_fam end', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'judicial_fam end'
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
        message = MIMEText('judicial_fam start', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'judicial_fam start'
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
        message = MIMEText('judicial_fam end', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'judicial_fam end'
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
        message = MIMEText('judicial_fam ERROR', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'judicial_fam ERROR'
        message['Subject'] = Header(subject, 'utf-8')
        try:
            smtpObj = smtplib.SMTP('vrh19.ucs.com')
            smtpObj.sendmail(sender, receivers, message.as_string())
            print('郵件傳送成功')
        except smtplib.SMTPException:
            print('Error: 無法傳送郵件')
def check(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select  type
    from UCS_Source  where  Engine='judicial_fam' and  status <> 'Done'  
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return list(c_src[0])[0]

def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num
def src_obs(server,username,password,database,fromtb,totb,check):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    
    select count(*) from judicial_fam_tstar_list where id not in (select id from judicial_fam_tstar) 
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]
def judicial(doc):
    judicial= []

    judicial.append({
        "Name":doc[0],
        "ID":doc[1],
        "announcement":doc[2],
        "post_date":doc[3],
        "register_no":doc[4],
        "keynote":doc[5],
        "Basis":doc[6],
        "Matters":doc[7],
        "update_date":doc[8],
        "note":doc[9],
        "flag":doc[10],
        "file_name":doc[11],
    })
    return judicial
def toSQL(docs, totb, server, database, username, password):
    import pyodbc
    #conn_cmd = 'DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password
    conn_cmd = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password
    with pyodbc.connect(conn_cmd) as cnxn:
        cnxn.autocommit = False
        with cnxn.cursor() as cursor:
            data_keys = ','.join(docs[0].keys())
            print(data_keys)
            data_symbols = ','.join(['?' for _ in range(len(docs[0].keys()))])
            print(data_symbols)
            insert_cmd = """INSERT INTO {} ({})
            VALUES ({})""".format(totb,data_keys,data_symbols )
            data_values = [tuple(doc.values()) for doc in docs]
            print(data_values)
            cursor.executemany( insert_cmd,data_values )
            cnxn.commit()


def dbfrom(server,username,password,database,fromtb,totb,check):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select * from judicial_fam_tstar_list where id not in (select id from judicial_fam_tstar)  order by ID
	offset 0 row fetch next 1 rows only
    
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def update(server,username,password,database,totb,note,ID,today,rowid,announcement,flag,post_date,register_no,keynote,Matters):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update [{totb}]
    set note = '{note}',update_date = '{today}',announcement='{announcement}',flag='{flag}',post_date='{post_date}',register_no='{register_no}',keynote='{keynote}',Matters='{Matters}'
    where id = '{ID}'
    
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()