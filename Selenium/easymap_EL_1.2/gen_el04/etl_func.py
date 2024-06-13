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
    from [{fromtb}] s1
    left join [dbo].[{totb}] s2
    on s1.addressi = s2.addressi and s1.eaid = s2.eaid and s1.pid = s2.pid and s1.LandNo=s2.LandNo
    where s2.LandNo is null and s1.type = '{entitytype}' and s1.Ld_District <> '' """    
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
    select s1.eaid,s1.pid,s1.addressi,s1.type,s1.Ld_District,s1.Ld_Lot,s1.LandNo,flg
    from [{fromtb}] s1
    left join [dbo].[{totb}] s2
    on s1.addressi = s2.addressi and s1.eaid = s2.eaid and s1.pid = s2.pid and s1.LandNo=s2.LandNo
    where s2.LandNo is null and s1.type = '{entitytype}' and s1.Ld_District <> ''
    order by s1.Ld_District desc--eaid,pid,addressi,LandNo
    offset 0 row fetch next 1 rows only"""    
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
    IF OBJECT_ID('EL04_TMP_{entitytype}', 'U') IS NULL    
    CREATE TABLE [dbo].[EL04_TMP_{entitytype}](
    [eaid] [int] NOT NULL,
    [pid] [nvarchar](50) NULL,
    [addressi] [int] NULL,
    [LandNo] [nvarchar](50) NULL,
    [purpose] [nvarchar](50) NULL,
    [appstatus] [nvarchar](5) NULL,
    [type] [nvarchar](50) NULL,
    [el_number] [nvarchar](50) NULL,
    [el_appdttm] [nvarchar](50) NULL,
    [app_result] [nvarchar](5) NULL,
    [insertdate] [datetime] NULL,
    [app_log] [nvarchar](800) NULL) 
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
            
def EL04_tmp_etl(doc):
    EL04_tmp_etl = []
    EL04_tmp_etl.append({
        "eaid":doc[0],
        "pid":doc[1],
        "addressi":doc[2],
        "Landno":doc[3],
        "purpose":doc[4],
        "appstatus":doc[5],
        "type":doc[6],
        "el_number":doc[7],
        "el_appdttm":doc[8],
        "app_result":doc[9],
        "insertdate":doc[10],
        "app_log":doc[11],
    })
    return EL04_tmp_etl

def APPOBS(server,username,password,database,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*) from [{totb}]
    where type = '{entitytype}'"""
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

def updatesql(server,username,password,database,servicetb,entitytype,EntityPhase,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    update [dbo].[{servicetb}]
    set validtrans = (SELECT count(*) FROM [UIS].[dbo].[{totb}] where app_result = 'Y')
    where entitytype = '{entitytype}' and entitytb = 'EL03DTL'
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

def updateMGM(server,username,password,database,MGtb,revobs,status,apply_dttm,entitytype,Machine,EntityPhase_next,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    update {MGtb} set validtrans = {revobs} ,status = '{status}',Statusdttm = '{apply_dttm}'
    where entitytype='{entitytype}' and Machine = '{Machine}' and EntityPhase = '{EntityPhase_next}' and EntityTb = '{totb}';
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    return script

def ERR_mail(obs):
    import smtplib
    from email.mime.text import MIMEText
    from email.header import Header
    sender = 'matthew5043@ucs.com'
    receivers = ['matthew5043@ucs.com'] # 接收郵件
    # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
    message = MIMEText('UCS EL 專案 Step 3 ERROR (10.90.0.72)', 'plain', 'utf-8')
    message['From'] = Header('UCSEL', 'utf-8') # 傳送者
    message['To'] = Header('matthew5043', 'utf-8') # 接收者
    subject = 'UCS EL 專案 Step 3 ERROR (10.90.0.72)'
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
    message = MIMEText('UCS EL 專案 Step 3 第二類謄本調閱 已完成(10.90.0.72)', 'plain', 'utf-8')
    message['From'] = Header('UCSEL', 'utf-8') # 傳送者
    message['To'] = Header('matthew5043', 'utf-8') # 接收者
    subject = 'UCS EL 專案 Step 3 第二類謄本調閱 已完成(10.90.0.72)'
    message['Subject'] = Header(subject, 'utf-8')
    try:
        smtpObj = smtplib.SMTP('vrh19.ucs.com')
        smtpObj.sendmail(sender, receivers, message.as_string())
        print('郵件傳送成功')
    except smtplib.SMTPException:
        print('Error: 無法傳送郵件')