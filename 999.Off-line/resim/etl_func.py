# -*- coding: utf-8 -*-
"""
Created on Mon May  9 18:28:59 2022

@author: admin
"""

#確認還沒有執行完的type


#寄送Mail 模板
def mail(obs):
    if obs ==0:
        import smtplib
        from email.mime.text import MIMEText
        from email.header import Header
        sender = 'matthew5043@ucs.com'
        receivers = ['matthew5043@ucs.com'] # 接收郵件
        # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
        message = MIMEText('resim end', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'resim end'
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
        message = MIMEText('resim start', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'resim start'
        message['Subject'] = Header(subject, 'utf-8')
        try:
            smtpObj = smtplib.SMTP('vrh19.ucs.com')
            smtpObj.sendmail(sender, receivers, message.as_string())
            print('郵件傳送成功')
        except smtplib.SMTPException:
            print('Error: 無法傳送郵件')

#寄送Mail 模板
def errormail(obs):
    if obs ==0:
        import smtplib
        from email.mime.text import MIMEText
        from email.header import Header
        sender = 'matthew5043@ucs.com'
        receivers = ['matthew5043@ucs.com'] # 接收郵件
        # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
        message = MIMEText('resim end', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'resim end'
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
        message = MIMEText('resim ERROR', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'resim ERROR'
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
    select (select count(*) from {fromtb} b
	left join [{totb}] r on b.ID = r.ID 
	where r.ID is null) 
	+
	(select count(*) from [{totb}] 
	where (select DATEDIFF(mm, update_date, getdate()))> = 3  and note  <> 'Y'  )

    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]
def resim(doc):
    resim= []

    resim.append({
        
        "Name":doc[0],
        "ID":doc[1],
        "member_type":doc[2],        
        "register_no":doc[3],
        "Receipt":doc[4],
        "Certificate":doc[5],
        "update_date":doc[6],
        "note":doc[7],
    })
    return resim
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

    
    select b.personi,b.ID,b.name,b.casei,b.type,b.c,b.m,b.age,b.flg,r.rowid,b.client_flg
    into #test
    from {fromtb} b
	left join [{totb}] r on b.ID = r.ID 
	where r.ID is null
	insert into #test
	select '0' as personi,ID,name,'0' as casei,'0' as type,'' as c,'' as m,'0' as age,'0' as flg,rowid,'0' as client_flg from [{totb}] 
	where (select DATEDIFF(mm, update_date, getdate()))> = 3  and note  <> 'Y' 
    select * from #test order by flg desc
    --order by personi
	offset 0 row fetch next 1 rows only
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def update(server,username,password,database,totb,note,ID,today,rowid,member_type,register_no,Receipt,Certificate):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update [{totb}]
    set note = '{note}',update_date = '{today}',member_type='{member_type}',register_no = '{register_no}',Receipt = '{Receipt}',Certificate = '{Certificate}'
    where id = '{ID}' and rowid = '{rowid}'
    
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