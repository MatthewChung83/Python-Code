def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num

def servicetime(server,username,password,database,servicetb,city,weekday):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select * from [{servicetb}] where county = '{city}' and weekday = {weekday}
    """
    cursor.execute(script)
    servicetime = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return servicetime

def src_obs(server,username,password,database,fromtb,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*)
    from [dbo].[{fromtb}] as a
    left join [dbo].[{totb}] as b
    on 
    a.[eaid]=b.[eaid] and a.[pid]=b.[pid] and a.[addressi]=b.[addressi] and 
    a.[LandNo]=b.[LandNo] and a.[el_number]=b.[el_number]
    where 
    a.type = '{entitytype}' and a.[app_result] = 'Y' and 
    a.[appstatus] = 'Y' and b.[el_number] is null
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]

def dbfrom(server,username,password,database,fromtb,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select a.[eaid], a.[pid], a.[addressi], a.[LandNo], a.[purpose], a.[type], a.[el_number], a.[el_appdttm]
    from [dbo].[{fromtb}] as a
    left join [dbo].[{totb}] as b
    on 
    a.[eaid]=b.[eaid] and a.[pid]=b.[pid] and a.[addressi]=b.[addressi] and 
    a.[LandNo]=b.[LandNo] and a.[el_number]=b.[el_number]
    where 
    a.type = '{entitytype}' and a.[app_result] = 'Y' and 
    a.[appstatus] = 'Y' and b.[el_number] is null
    order by eaid,pid,addressi,LandNo
    offset 0 row fetch next 1 rows only
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

# 確認pk資料是否已在 DB table 中
def createtmp(server,username,password,database,fromtb,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    IF OBJECT_ID('EL05_TMP_{entitytype}', 'U') IS NULL    
    CREATE TABLE [dbo].[EL05_TMP_{entitytype}](
    [eaid] [int] NOT NULL,
    [pid] [nvarchar](50) NULL,
    [addressi] [int] NULL,
    [LandNo] [nvarchar](50) NULL,
    [purpose] [nvarchar](50) NULL,
    [type]  [nvarchar](50) NULL,
    [number_src] [nvarchar](50) NULL,
    [el_number] [nvarchar](100) Null,
    [pages] [nvarchar](5) Null,
    [app_fee] [nvarchar](5) Null,
    [file_save] [nvarchar](300) Null,
    [el_appdttm] [nvarchar](50) NULL,
    [rev_appdttm] [datetime] NULL,
    [note] [nvarchar](500) NULL,
    [insertdate] [datetime] NULL,
    [rev_log] [nvarchar](800) NULL)
    ON [PRIMARY]"""
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    
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
            
def EL05_tmp_etl(doc):
    EL05_tmp_etl = []
    EL05_tmp_etl.append({
        "eaid":doc[0],
        "pid":doc[1],
        "addressi":doc[2],
        "LandNo":doc[3],
        "purpose":doc[4],
        "type":doc[5],
        "number_src":doc[6],
        "el_number":doc[7],
        "pages":doc[8],
        "app_fee":doc[9],
        "file_save":doc[10],
        "el_appdttm":doc[11],
        "rev_appdttm":doc[12],
        "note":doc[13],
        "insertdate":doc[14],
        "rev_log":doc[15],
    })
    return EL05_tmp_etl

def REVOBS(server,username,password,database,fromtb,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*) from 
    (
    select distinct eaid,pid,addressi,landno,el_number
    from [{totb}]
    where type = '{entitytype}'
    ) as a
    """    
    cursor.execute(script)
    REVOBS = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(REVOBS[0])[0]

def APPOBS(server,username,password,database,fromtb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*) from 
    (
    select distinct eaid,pid,addressi,landno,el_number
    from [{fromtb}] 
    where app_result = 'Y' and type = '{entitytype}'
    ) as a
    """
    cursor.execute(script)
    REVOBS = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(REVOBS[0])[0]

def APPSUCCESSOBS(server,username,password,database,fromtb,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*) from [{totb}]
    where type = '{entitytype}' and app_result = 'Y'"""    
    cursor.execute(script)
    DTLOBS = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(DTLOBS[0])[0]

def updateMGM(server,username,password,database,MGtb,revobs,status,apply_dttm,entitytype,Machine,EntityPhase_last,fromtb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    update {MGtb} set validtrans = {revobs} ,status = '{status}',Statusdttm = '{apply_dttm}'
    where entitytype='{entitytype}' and Machine = '{Machine}' and EntityPhase = '{EntityPhase_last}' and EntityTb = '{fromtb}';
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    return script

def insertMG(server,username,password,database,MGtb,entitytype,Machine,EntityPhase_next,totb,appsuccessobs,status,apply_dttm,EntityPath_next):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    if exists(SELECT * from {MGtb} WHERE entitytype='{entitytype}' and Machine = '{Machine}' and EntityPhase = '{EntityPhase_next}' and EntityTb = '{totb}' )
        begin
        update {MGtb} set EntityObs = {appsuccessobs},status = '{status}', Statusdttm = '{apply_dttm}' 
        WHERE entitytype='{entitytype}' and Machine = '{Machine}' and EntityPhase = '{EntityPhase_next}' and EntityTb = '{totb}'
        end    
    else
        begin
        insert into {MGtb} values('{entitytype}','{Machine}','{EntityPhase_next}','{EntityPath_next}',0,null,'{status}','{apply_dttm}','{apply_dttm}','{totb}')
        end
    """    
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    return script

def existfile(file_path):    
    import os    
    if os.path.isfile(file_path):
        return True
    else:
        return False
    
def ERR_mail(obs):
    import smtplib
    from email.mime.text import MIMEText
    from email.header import Header
    sender = 'matthew5043@ucs.com'
    receivers = ['matthew5043@ucs.com'] # 接收郵件
    # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
    message = MIMEText('UCS EL 專案 Step 4 第二類謄本上傳圖檔 ERROR (10.90.0.72)', 'plain', 'utf-8')
    message['From'] = Header('UCSEL', 'utf-8') # 傳送者
    message['To'] = Header('matthew5043', 'utf-8') # 接收者
    subject = 'UCS EL 專案 Step 4 第二類謄本上傳圖檔 ERROR (10.90.0.72)'
    message['Subject'] = Header(subject, 'utf-8')
    try:
        smtpObj = smtplib.SMTP('vrh19.ucs.com')
        smtpObj.sendmail(sender, receivers, message.as_string())
        print('郵件傳送成功')
    except smtplib.SMTPException:
        print('Error: 無法傳送郵件')

def SUCC_mail(obs):
    #Send Mail        
    import smtplib
    from email.mime.text import MIMEText
    from email.header import Header
    sender = 'matthew5043@ucs.com'
    receivers = ['matthew5043@ucs.com'] # 接收郵件
    # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
    message = MIMEText('UCS EL 專案 Step 4 第二類謄本下載/上傳圖檔 已完成(10.90.0.72)', 'plain', 'utf-8')
    message['From'] = Header('UCSEL', 'utf-8') # 傳送者
    message['To'] = Header('matthew5043', 'utf-8') # 接收者
    subject = 'UCS EL 專案 Step 4 第二類謄本下載/上傳圖檔 已完成(10.90.0.72)'
    message['Subject'] = Header(subject, 'utf-8')
    try:
        smtpObj = smtplib.SMTP('vrh19.ucs.com')
        smtpObj.sendmail(sender, receivers, message.as_string())
        print('郵件傳送成功')
    except smtplib.SMTPException:
        print('Error: 無法傳送郵件')
def ERR_mail_step2(obs):
    import smtplib
    from email.mime.text import MIMEText
    from email.header import Header
    sender = 'matthew5043@ucs.com'
    receivers = ['matthew5043@ucs.com'] # 接收郵件
    # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
    message = MIMEText('UCS EL 專案 Step 4 第二類謄本下載圖檔 ERROR (10.90.0.72)', 'plain', 'utf-8')
    message['From'] = Header('UCSEL', 'utf-8') # 傳送者
    message['To'] = Header('matthew5043', 'utf-8') # 接收者
    subject = 'UCS EL 專案 Step 4 第二類謄本下載圖檔 ERROR (10.90.0.72)'
    message['Subject'] = Header(subject, 'utf-8')
    try:
        smtpObj = smtplib.SMTP('vrh19.ucs.com')
        smtpObj.sendmail(sender, receivers, message.as_string())
        print('郵件傳送成功')
    except smtplib.SMTPException:
        print('Error: 無法傳送郵件')