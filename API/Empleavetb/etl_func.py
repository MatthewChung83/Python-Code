# -*- coding: utf-8 -*-
"""
Created on Mon May  9 18:28:59 2022

@author: admin
"""



def dbfrom(server,username,password,database,totb,EMPLOYEEID,SYS_VIEWID,STARTDATE,STARTTIME,ENDDATE,ENDTIME):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select * from {totb} where  [EMPLOYEEID] = '{EMPLOYEEID}' and [SYS_VIEWID]='{SYS_VIEWID}' and [STARTDATE] = '{STARTDATE}'
    and [STARTTIME] = '{STARTTIME}' and [ENDDATE] = '{ENDDATE}' and [ENDTIME] = '{ENDTIME}'
    
    order by rowid desc
	offset 0 row fetch next 1 rows only
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def update(server,username,password,database,totb,EMPLOYEEID,SYS_FLOWFORMSTATUS,SYS_VIEWID,update_date,STARTDATE,STARTTIME,ENDDATE,ENDTIME):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update [{totb}]
    set [SYS_FLOWFORMSTATUS] = '{SYS_FLOWFORMSTATUS}',[update_date]='{update_date}',[STARTDATE] = '{STARTDATE}',[STARTTIME] = '{STARTTIME}',[ENDDATE] = '{ENDDATE}',[ENDTIME] = '{ENDTIME}'
    where [EMPLOYEEID] = '{EMPLOYEEID}' and [SYS_VIEWID]='{SYS_VIEWID}' and [STARTDATE] = '{STARTDATE}'
    and [STARTTIME] = '{STARTTIME}' and [ENDDATE] = '{ENDDATE}' and [ENDTIME] = '{ENDTIME}'
    
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    
def delete(server,username,password,database,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    delete from  [{totb}]
    
    """
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

def mail():
    import smtplib
    from email.mime.text import MIMEText
    from email.header import Header
    sender = 'matthew5043@ucs.com'
    receivers = ['matthew5043@ucs.com'] # 接收郵件
    # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
    message = MIMEText('empleavetb end', 'plain', 'utf-8')
    message['From'] = Header('matthew5043', 'utf-8') # 傳送者
    message['To'] = Header('matthew5043', 'utf-8') # 接收者
    subject = 'empleavetb end'
    message['Subject'] = Header(subject, 'utf-8')
    try:
        smtpObj = smtplib.SMTP('vrh19.ucs.com')
        smtpObj.sendmail(sender, receivers, message.as_string())
        print('郵件傳送成功')
    except smtplib.SMTPException:
        print('Error: 無法傳送郵件')