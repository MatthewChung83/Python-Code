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
        message = MIMEText('etwmain start', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'etwmain start'
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
        message = MIMEText('etwmain end', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'etwmain end'
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
        message = MIMEText('etwmain ERROR', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'etwmain ERROR'
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

def dbfrom(server,username,password,database,fromtb,fromtb1,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database,autocommit=True)
    cursor = conn.cursor()
    
    script = f"""
    insert into treasure.skiptrace.dbo.Legal_G_List(query_id,masterclienti,type,id,institution)
    select distinct fr.query_id,frl.masterclienti,frl.query_type ,fr.query_val,replace(replace(replace(replace(fr.query_company,' ',''),CHAR(9),''),CHAR(10),''),CHAR(13),'')
    from finedb.dbo.fr_query_list_detail fr
    left join finedb.dbo.fr_query_list frl on frl.uuid = fr.query_id
    left join treasure.skiptrace.dbo.Legal_G_List LG on frl.uuid = LG.query_id  and LG.institution = replace(replace(replace(replace(fr.query_company,' ',''),CHAR(9),''),CHAR(10),''),CHAR(13),'')
    left join FRDB.dbo.persontb p on p.id = fr.query_val
    left join FRDB.dbo.casepersontb cp on p.personi = cp.personi 
    left join FRDB.dbo.casetb c on cp.casei = c.casei
    where c.statusi not in (5,9,10,17,18,32,36,64,87,92,101,141,147,149,150)  and p.condition <> '去世' 
    and lg.institution is null 
    order by fr.query_val
    
	declare @0 int,@1 int,@2 int,@3 int,@4 int
    select @0 = min(rowid) from treasure.skiptrace.dbo.Legal_G_List where status is null
    select @1 = @0 + count(*)/4 from treasure.skiptrace.dbo.Legal_G_List 
    select @2 = @1 + count(*)/4 from treasure.skiptrace.dbo.Legal_G_List 
    select @3 = @2 + count(*)/4 from treasure.skiptrace.dbo.Legal_G_List 
    select @4 = MAX(rowid) from treasure.skiptrace.dbo.Legal_G_List where status is null 
    --select @0,@1,@2,@3,@4
    update treasure.skiptrace.dbo.Legal_G_List set list_type = '01' where rowid >=@0 and rowid <=@1 and status is null and list_type is null
    update treasure.skiptrace.dbo.Legal_G_List set list_type = '02' where rowid >@1 and rowid <=@2 and status is null and list_type is null
    update treasure.skiptrace.dbo.Legal_G_List set list_type = '03' where rowid >@2 and rowid <=@3 and status is null and list_type is null
    update treasure.skiptrace.dbo.Legal_G_List set list_type = '04' where rowid >@3 and rowid <=@4 and status is null and list_type is null
    
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()

def Legal_G_List(doc):
    Legal_G_List= []

    Legal_G_List.append({
        "query_id":doc[0],
        "masterclienti":doc[1],
        "type":doc[2],
        "id":doc[3],
        "Institution":doc[4],
        "Third_Company":doc[5],
        "Third_Name":doc[6],
        "Third_num":doc[7],
        "Business_address":doc[8],
        "Business_status":doc[9],
        "Capital_amount":doc[10],
        "Tissue_type":doc[11],
        "Date_of_establishment":doc[12],
        "Register_business_items":doc[13],
        "Note":doc[14],
        "Casei":doc[15],
        "Legal_status":doc[16],
        "Court":doc[17],
        "Insert_date":doc[18],
        "status":doc[19]
        
    })
    return Legal_G_List

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



def src_obs_U(server,username,password,database,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database,autocommit=True)
    cursor = conn.cursor()
    script = f"""
    
    select  count(distinct Institution)
    from treasure.skiptrace.dbo.Legal_G_List
    where status is null and list_type = '04'
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]

def dbfrom_U(server,username,password,database,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database,autocommit=True)
    cursor = conn.cursor()
    
    script = f"""
    select  distinct Institution
    from treasure.skiptrace.dbo.Legal_G_List
    where status is null and list_type = '04'
    order by Institution
	offset 0 row fetch next 1 rows only
    
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src
def update(server,username,password,database,totb,query_company,Thrid_Company,Thrid_Name,Thrid_Num,Business_address,Business_status,Capital_amount,Tissue_type,Date_of_establishment,Register_business_items,today,status):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database,autocommit=True)
    cursor = conn.cursor()
        
    script = f"""
    update treasure.skiptrace.dbo.Legal_G_List
    set Third_Company='{Thrid_Company}',Third_Name='{Thrid_Name}',Third_Num='{Thrid_Num}',Business_address='{Business_address}',Business_status='{Business_status}',Capital_amount='{Capital_amount}',Tissue_type='{Tissue_type}'
    ,Date_of_establishment='{Date_of_establishment}',Register_business_items='{Register_business_items}',Insert_date='{today}',status='{status}'
    where Institution = '{query_company}' and status is null
    
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()