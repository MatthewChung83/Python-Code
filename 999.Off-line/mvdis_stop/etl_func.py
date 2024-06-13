# -*- coding: utf-8 -*-
"""
Created on Mon May  9 18:43:32 2022

@author: admin
"""

def num(obs,server,username,password,database,fromtb,totb,check):
    obs = src_obs(server,username,password,database,fromtb,totb,check)
    if obs > 0 :
        import pymssql
        conn = pymssql.connect(server=server, user=username, password=password, database = database)
        cursor = conn.cursor()
        
        script = f"""
        update UCS_Source
        set F_OBS = '{obs}'
        where Engine = 'dead' and type = '{check}'
        """
        cursor.execute(script)
        conn.commit()
        cursor.close()
        conn.close()
    elif obs ==0 :
        import pymssql
        conn = pymssql.connect(server=server, user=username, password=password, database = database)
        cursor = conn.cursor()
        
        script = f"""
        update UCS_Source
        set status = 'Done'
        where Engine = 'dead' and type = '{check}'
        """
        cursor.execute(script)
        conn.commit()
        cursor.close()
        conn.close()
def check(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select  type
    from UCS_Source  where  Engine='dead' and  status <> 'Done'  
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
    select (select count(*) from base_case b
	left join dead r on b.ID = r.ID 
	where r.ID is null  )
	+
	(select count(*) from base_case b
	left join dead r on b.ID = r.ID 
	where (select DATEDIFF(mm, update_date, getdate()))> = 3  and note  <> 'Y'  )
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]
def dead(doc):
    dead= []

    dead.append({
        #"rowid":doc[0],
        "Name":doc[0],
        "ID":doc[1],
        "birthday":doc[2],
        "update_date":doc[3],
        "note":doc[4],
    })
    return dead
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


def dbfrom(server,username,password,database,fromtb,totb,check):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""

    select b.*,r.rowid
	into #test
	from base_case (nolock) b
	left join dead (nolock) r on b.ID = r.ID 
	where r.ID is null 
	insert into #test
	select b.*,r.rowid from base_case (nolock) b
	left join dead (nolock) r on b.ID = r.ID 
	where (select DATEDIFF(mm, update_date, getdate()))> = 3  and note  <> 'Y' 
	select * from #test 
    order by personi
	offset 0 row fetch next 1 rows only
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src


def CAPTCHAOCR(Apurl,imgf,imgp):
    import requests
    url = Apurl
    
    files = {
        "image_file":(imgf,open(imgp,"rb"),'image/jpeg')
    }
    response = requests.request("POST", url, files=files)
    return response.text
def mail(obs):
    if obs ==0:
        import smtplib
        from email.mime.text import MIMEText
        from email.header import Header
        sender = 'matthew5043@ucs.com'
        receivers = ['matthew5043@ucs.com'] # 接收郵件
        # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
        message = MIMEText('dead end', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'dead end'
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
        message = MIMEText('dead start', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'dead start'
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
        message = MIMEText('dead end', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'dead end'
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
        message = MIMEText('dead ERROR', 'plain', 'utf-8')
        message['From'] = Header('matthew5043', 'utf-8') # 傳送者
        message['To'] = Header('matthew5043', 'utf-8') # 接收者
        subject = 'dead ERROR'
        message['Subject'] = Header(subject, 'utf-8')
        try:
            smtpObj = smtplib.SMTP('vrh19.ucs.com')
            smtpObj.sendmail(sender, receivers, message.as_string())
            print('郵件傳送成功')
        except smtplib.SMTPException:
            print('Error: 無法傳送郵件')

def capcha_resp(url,captchaImg,imgp,q,ID,bir):
# url:Post網址
# captchaImg:辨識碼網址
# imgp:辨識碼本機儲存路徑
# q:辨識迭代次數
# ID:Post-form data的ID欄位
# birthday:Post-form data的出生日期欄位

    from bs4 import BeautifulSoup
    import requests
    import ddddocr
    
    while q < 100:
        # 取辨識碼並儲存至Local
        session = requests.session()
        respimg = session.get(captchaImg)
        img = open(imgp, 'wb')
        img.write(respimg.content)
        img.close()
        
        # 執行辨識作業
        ocr = ddddocr.DdddOcr()
        with open(imgp, 'rb') as f:
            img_bytes = f.read()
            res = ocr.classification(img_bytes)
        
        # post API then get result
        data = {
            'method': 'queryCheck',
            'queryType': '1',
            'companyNo' : '',
            'idNo': f'{ID}',
            'birthday': f'{bir}',
            'validateStr': f'{res}',
        }
        
        resp = session.post(url,data=data)
        soup = BeautifulSoup(resp.text,"lxml")
        
        # 判斷辨識碼是否正確辨別
        if soup.find_all('span',string = '驗證碼輸入錯誤'):
            session.close()
            q =+ 1
        else:
            session.close()
            result = soup
            return result
            q = 100
            
def update(server,username,password,database,totb,note,ID,today,rowid):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update [{totb}]
    set note = '{note}',update_date = '{today}'
    where id = '{ID}' and rowid = '{rowid}'
    
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()