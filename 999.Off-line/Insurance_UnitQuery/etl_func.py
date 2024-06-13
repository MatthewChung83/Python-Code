def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num
def check(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select  type
    from UCS_Source  where  Engine='Insurance_UnitQuery' and  status <> 'Done'  
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return list(c_src[0])[0]
def mail(obs):
    if obs ==0:
        import smtplib
        from email.mime.text import MIMEText
        from email.header import Header
        sender = 'matthew5043@ucs.com'
        receivers = ['matthew5043@ucs.com'] # 接收郵件
        # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
        message = MIMEText('Insurance_UnitQuery end', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'Insurance_UnitQuery end'
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
        message = MIMEText('Insurance_UnitQuery start', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'Insurance_UnitQuery start'
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
        message = MIMEText('Insurance_UnitQuery end', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'Insurance_UnitQuery end'
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
        message = MIMEText('Insurance_UnitQuery ERROR', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'Insurance_UnitQuery ERROR'
        message['Subject'] = Header(subject, 'utf-8')
        try:
            smtpObj = smtplib.SMTP('vrh19.ucs.com')
            smtpObj.sendmail(sender, receivers, message.as_string())
            print('郵件傳送成功')
        except smtplib.SMTPException:
            print('Error: 無法傳送郵件')
def dbfrom(server,username,password,database,totb1,totb2,check):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""    
    
    select b.personi,b.ID,b.name,b.casei,b.type,b.c,b.m,b.age,b.flg,r.rowid,b.client_flg
	into #test
	from [{totb1}] (nolock) b
	left join [{totb2}] (nolock) r on b.ID = r.ID 
	where r.ID is null --and flg = '{check}'
	insert into #test
	select '0' as personi,ID,name,'0' as casei,'0' as type,'' as c,'' as m,'0' as age,'0' as flg,rowid,'0' as client_flg from [{totb2}] 
	where (select DATEDIFF(mm, update_date, getdate()))> = 3  and status  <> 'Y' 
	select * from #test order by flg desc
    --order by personi
	offset 0 row fetch next 1 rows only
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src
    
def src_obs(server,username,password,database,totb1,totb2,check):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""

    select (select count(*) from base_case b
	left join Insurance_UnitQuery r on b.ID = r.ID 
	where r.ID is null  )--and flg= '{check}')
	+
	(select count(*) from base_case b
	left join Insurance_UnitQuery r on b.ID = r.ID 
	where (select DATEDIFF(mm, update_date, getdate()))> = 3  and Status  <> 'Y'  )
    
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]

def toSQL(docs, totb2, server, database, username, password):
    import pyodbc
    #conn_cmd = 'DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password
    conn_cmd = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password
    with pyodbc.connect(conn_cmd) as cnxn:
        cnxn.autocommit = False
        with cnxn.cursor() as cursor:
            data_keys = ','.join(docs[0].keys())
            data_symbols = ','.join(['?' for _ in range(len(docs[0].keys()))])
            insert_cmd = """INSERT INTO {} ({})
            VALUES ({})""".format(totb2,data_keys,data_symbols )
            data_values = [tuple(doc.values()) for doc in docs]
            cursor.executemany( insert_cmd,data_values )
            cnxn.commit()

def indata(doc):
    etl = []
    etl.append({

        'Name':doc[0],
        'ID':doc[1],
        'Response':doc[2],
        'update_date':doc[3],
        'Status':doc[4],
        
    })
    return etl
 
def indetail(doc):
    etl = []
    for debtor in doc[7]:
        etl.append({
            "judicial_no" : doc[0],
            "recipient": doc[1],
            "publishing_date": doc[2],
            "judicial_doc_date": doc[3],
            "judicial_doc_no": doc[4],
            "judicial_doc_subject": doc[5],
            "debtors": doc[6],
            "debtor": debtor,
            "insertdate": doc[8],
            })
    return etl

def update(server,username,password,database,totb2,status,ID,today,rowid,response):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update [{totb2}]
    set status = '{status}',update_date = '{today}',response = '{response}'
    where id = '{ID}' and rowid = '{rowid}'
    
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    
def exit_obs(server,username,password,database,totb2):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count (distinct id)
    from [{totb2}]
    where  update_date > = convert(varchar(10),getdate(),111)
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]