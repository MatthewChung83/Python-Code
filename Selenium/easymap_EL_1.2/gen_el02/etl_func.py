def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num

def src_obs(server,username,password,database,fromtb,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    if '{entitytype}' like 'AA_%' or '{entitytype}' like 'AM_%'
	begin
    select count(*) from [{fromtb}](nolock)
    where type = '{entitytype}' and eaid not in (select eaid from {totb}(nolock)) and city is not null
    end"""    
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

    select [eaid],[pid],[addressi],[casei],[city],[town],[street],[sec],[lane],[lane2],[no],[subno],[floor],[subfloor],[type] 
    from [{fromtb}](nolock) where type = '{entitytype}' and eaid not in (select eaid from {totb}(nolock)) --and city is not null
    order by eaid
    offset 0 row fetch next 1 rows only

    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

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

def EL02_etl(doc):
    ### auction_info_tb
    EL02_result = []

    EL02_result.append({
        "eaid":doc[0],
        "pid":doc[1],
        "addressi":doc[2],
        "casei":doc[3],
        "City":doc[4],
        "Town":doc[5],
        "Street":doc[6],
        "Lane":doc[7],
        "alley":doc[8],
        "No":doc[9],
        "Floor":doc[10],
        "type":doc[11],
        "result":doc[12],
        "datatime":doc[13],
    })
    return EL02_result

def updatesql(server,username,password,database,validtrans,actiondttm,Statusdttm,status,entitytype,Machine,EntityPhase):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    if '{entitytype}' like 'AA_%' or '{entitytype}' like 'AM_%'
	begin
    update EL_Entity_Managerment 
    set validtrans = {validtrans} ,actiondttm = '{actiondttm}', Statusdttm = '{Statusdttm}', status = '{status}'
    WHERE entitytype='{entitytype}' and Machine = '{Machine}' and EntityPhase = '{EntityPhase}'
    end;
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
    message = MIMEText('UCS EL 專案 Step 1 ERROR (10.90.0.72)', 'plain', 'utf-8')
    message['From'] = Header('wbt', 'utf-8') # 傳送者
    message['To'] = Header('matthew5043', 'utf-8') # 接收者
    subject = 'UCS EL 專案 Step 1 ERROR (10.90.0.72)'
    message['Subject'] = Header(subject, 'utf-8')
    try:
        smtpObj = smtplib.SMTP('vrh19.ucs.com')
        smtpObj.sendmail(sender, receivers, message.as_string())
        print('郵件傳送成功')
    except smtplib.SMTPException:
        print('Error: 無法傳送郵件')

def SUCC_mail(obs_wbt):
    #Send Mail        
    import smtplib
    from email.mime.text import MIMEText
    from email.header import Header
    sender = 'matthew5043@ucs.com'
    receivers = ['matthew5043@ucs.com'] # 接收郵件
    # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
    message = MIMEText('UCS EL 專案 Step 1 地址驗證 已完成(10.90.0.72)', 'plain', 'utf-8')
    message['From'] = Header('wbt', 'utf-8') # 傳送者
    message['To'] = Header('matthew5043', 'utf-8') # 接收者
    subject = 'UCS EL 專案 Step 1 地址驗證 已完成(10.90.0.72)'
    message['Subject'] = Header(subject, 'utf-8')
    try:
        smtpObj = smtplib.SMTP('vrh19.ucs.com')
        smtpObj.sendmail(sender, receivers, message.as_string())
        print('郵件傳送成功')
    except smtplib.SMTPException:
        print('Error: 無法傳送郵件')
        
def new_collection(server,username,password,database,entitytype,fromtb,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    insert into {totb} (eaid,pid,addressi,casei,city,town,street,lane,alley,no,floor,result,datatime,type)
    select  e0.eaid,e0.pid,e0.addressi,e0.casei,e0.City,e0.town,e0.Street+e0.sec as street,e0.lane,e0.Lane2 as alley,e0.No,e0.Floor,
    case when ai.city is null then 'N' else 'Y' end as result,
    convert(datetime,getdate()) as [datatime],
    e0.type
     from {fromtb} e0
    left join address_index ai on e0.City = ai.city and e0.Town = ai.town and e0.Street+e0.Sec = ai.road
    where eaid not in (select eaid from {totb}) and e0.type = '{entitytype}'
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    



