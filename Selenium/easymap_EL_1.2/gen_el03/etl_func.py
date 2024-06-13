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
    select count(*) from [{fromtb}]
    where result = 'Y' and type = '{entitytype}' and eaid not in (select eaid from {totb})
    end
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]

def DTLOBS(server,username,password,database,fromtb,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    if '{entitytype}' like 'AA_%' or '{entitytype}' like 'AM_%'
	begin
    select count(*) from [EL03DTL]
    where type = '{entitytype}'
    end"""    
    cursor.execute(script)
    DTLOBS = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(DTLOBS[0])[0]

def dbfrom(server,username,password,database,fromtb,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    if '{entitytype}' like 'AA_%' or '{entitytype}' like 'AM_%'
	begin
    select * from [{fromtb}] where result = 'Y' and type = '{entitytype}' and eaid not in (select eaid from {totb})
    order by eaid
    offset 0 row fetch next 1 rows only
    end
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def toSQL(docs, totb, server, database, username, password):
    import pyodbc
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
            
def EL03_etl(doc):
    EL03_result = []
    EL03_result.append({
        "eaid":doc[0],
        "pid":doc[1],
        "addressi":doc[2],
        "casei":doc[3],
        "type":doc[4],
        
        "Bd_District":doc[5],
        "Bd_GeoOffice":doc[6],
        "Bd_Lot":doc[7],
        "BuildNo":doc[8],
        "Bd_AreaSq":doc[9],
        "Bd_FloorNo":doc[10],
        "Bd_ByFloor":doc[11],
        "Bd_CompletionDate":doc[12],
        "Bd_Purpose":doc[13],
        
        "Ld_District":doc[14],
        "Ld_GeoOffice":doc[15],
        "Ld_Lot":doc[16],
        "LandNo":doc[17],
        "LandNo_f":doc[18],
        "LandNo_b":doc[19],
        "Ld_AreaSq":doc[20],
        "Ld_Value":doc[21],
        "Ld_Price":doc[22],
        
        "result":doc[23],
        "insertdate":doc[24],
        
        "bd_log":doc[25],
        "ld_log":doc[26], 
        "bd_landno":doc[27],
    })
    return EL03_result

def EL03_DTL_etl(doc):
    EL03_DTL_result = []
    EL03_DTL_result.append({
        "eaid":doc[0],
        "pid":doc[1],
        "addressi":doc[2],
        "type":doc[3],
        "Ld_District":doc[4],
        "Ld_Lot":doc[5],
        "LandNo":doc[6],
        "flg":doc[7],        
        "insertdate":doc[8],
    })
    return EL03_DTL_result

def updateMG_Main(server,username,password,database,validtrans,actiondttm,Statusdttm,status,entitytype,Machine,EntityPhase,MGtb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    if '{entitytype}' like 'AA_%' or '{entitytype}' like 'AM_%'
	begin
    update {MGtb} 
    set validtrans = {validtrans} ,actiondttm = '{actiondttm}', Statusdttm = '{Statusdttm}', status = '{status}'
    WHERE entitytype='{entitytype}' and Machine = '{Machine}' and EntityPhase = '{EntityPhase}' and EntityTb = 'EL03'
    end;
    """    
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    return script

def updateMG_Next(server,username,password,database,actiondttm,Statusdttm,status,entitytype,Machine,EntityPhase_next,EntityPath_next,DTLOBS,MGtb,todtltb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    if '{entitytype}' like 'AA_%' or '{entitytype}' like 'AM_%'
    	begin
    if exists(SELECT * from {MGtb} WHERE entitytype='{entitytype}' and Machine = '{Machine}' and EntityPhase = '{EntityPhase_next}' and EntityTb = '{todtltb}' )
        begin
        update {MGtb} set EntityObs = {DTLOBS},actiondttm = '{actiondttm}',status = '{status}', Statusdttm = '{Statusdttm}' WHERE entitytype='{entitytype}' and Machine = '{Machine}' and EntityPhase = '{EntityPhase_next}' and EntityTb = '{todtltb}'
        end    
    else
        begin
        insert into {MGtb} values('{entitytype}','{Machine}','{EntityPhase_next}','{EntityPath_next}',{DTLOBS},null,'{status}','{actiondttm}','{Statusdttm}','{todtltb}')
        end
        end
    """    
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    return script

def check_obs(server,username,password,database,fromtb,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    if '{entitytype}' like 'AA_%' or '{entitytype}' like 'AM_%'
    begin
    select count(*) from [{totb}]
    where type = '{entitytype}'
    end"""    
    cursor.execute(script)
    obs = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return list(obs[0])[0]

def landno_split(landno,landno_f,landno_b,bd_landno):  
    if landno == '':
        landno_str = ''
    elif landno_b == '0000':
        landno_str = str(int(landno_f))
    else:
        landno_str = str(int(landno_f)) + '-' + str(int(landno_b))
    
    bd_landno_sp = bd_landno.split(',',-1)
    bd_landno_sp.append(landno_str)
    bd_landno_tmp = list(set(filter(None,bd_landno_sp)))
    landno_list = []
    flg= ''
    for t in range(len(bd_landno_tmp)):
        if landno == '':
            if '-' in bd_landno_tmp[t]:
                pass
            else:
                bd_landno_tmp[t] = bd_landno_tmp[t] + '-' + '0000'
        else:
            if '-' in bd_landno_tmp[t]:
                pass
            elif landno_b == '0000':
                bd_landno_tmp[t] = bd_landno_tmp[t]
            else:
                bd_landno_tmp[t] = bd_landno_tmp[t] + '-' + str(int(landno_b))
        
        if bd_landno_tmp[t] == landno_str or bd_landno == '':
            flg = 'same as landno'
        else:
            flg = 'new by buildno'
        
        text = (bd_landno_tmp[t],flg)
        landno_list.append(text)
    
    return landno_list

def ERR_mail(obs):
    import smtplib
    from email.mime.text import MIMEText
    from email.header import Header
    sender = 'matthew5043@ucs.com'
    receivers = ['matthew5043@ucs.com'] # 接收郵件
    # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
    message = MIMEText('UCS EL 專案 Step 2 ERROR (10.90.0.72)', 'plain', 'utf-8')
    message['From'] = Header('UCSEL', 'utf-8') # 傳送者
    message['To'] = Header('matthew5043', 'utf-8') # 接收者
    subject = 'UCS EL 專案 Step 2 ERROR (10.90.0.72)'
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
    message = MIMEText('UCS EL 專案 Step 2 地址轉地號 已完成(10.90.0.72)', 'plain', 'utf-8')
    message['From'] = Header('UCSEL', 'utf-8') # 傳送者
    message['To'] = Header('matthew5043', 'utf-8') # 接收者
    subject = 'UCS EL 專案 Step 2 地址轉地號 已完成(10.90.0.72)'
    message['Subject'] = Header(subject, 'utf-8')
    try:
        smtpObj = smtplib.SMTP('vrh19.ucs.com')
        smtpObj.sendmail(sender, receivers, message.as_string())
        print('郵件傳送成功')
    except smtplib.SMTPException:
        print('Error: 無法傳送郵件')