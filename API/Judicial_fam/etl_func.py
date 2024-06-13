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

def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num
def src_obs(server,username,password,database,fromtb,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    
    select ( select count(*) from base_case_allID b
            left join judicial_fam r on b.ID = r.ID 
            where r.ID is null )
   --+
   -- (select count(*) from judicial_fam 
   -- where (select DATEDIFF(mm, update_date, getdate()))> = 6 )  
	 
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
        "note2":doc[11],
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
            data_symbols = ','.join(['?' for _ in range(len(docs[0].keys()))])
            insert_cmd = """INSERT INTO {} ({})
            VALUES ({})""".format(totb,data_keys,data_symbols )
            data_values = [tuple(doc.values()) for doc in docs]
            cursor.executemany( insert_cmd,data_values )
            cnxn.commit()


def dbfrom(server,username,password,database,fromtb,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select distinct b.personi,b.ID,b.name,b.casei,b.type,b.c,b.m,b.age,b.flg,r.rowid
    into #test
    from {fromtb} b
	left join [{totb}] r on b.ID = r.ID 
	where r.ID is null
    
    
    insert into #test
    select distinct 0 as personi,ID,name,0 as casei,0 as type,0 as c,0 as m,0 as age,'' as flg,rowid
    from [{totb}]
    where (select DATEDIFF(mm, update_date, getdate()))> = 3
	
    select * from #test order by flg desc 
	--offset 0 row fetch next 1 rows only
    
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def delete(server,username,password,database,totb,note,ID,rowid,register_no,item):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    delete from [{totb}]
    where id = '{ID}'  --and register_no = '{register_no}' and flag = '{item}'
    
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
def exit_obs(server,username,password,database,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count (distinct ID)
    from [{totb}]
    where  update_date > = convert(varchar(10),getdate(),111)
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]