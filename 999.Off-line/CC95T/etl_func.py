# -*- coding: utf-8 -*-
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
    where r.ID is null and b.C is not null) 
    +
    (select count(*) from [{totb}] 
    where (select DATEDIFF(mm, updatetime, getdate()))> = 3 and status <> 'Y' )  
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]

def dbfrom(server,username,password,database,fromtb,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select b.personi,b.ID,b.name,convert(varchar(10),b.C,111) as C,b.m,b.flg
    into #test
    from {fromtb} b
    left join [{totb}] r on b.ID = r.ID 
    where r.ID is null and b.C is not null
    insert into #test
    select r.personi,r.ID,r.name,convert(varchar(10),r.M,111) as C,r.m,'' as flg  
    from  [{totb}] r 
    where (select DATEDIFF(mm, r.updatetime, getdate()))> = 3 and r.status <> 'Y' 
    select * from #test order by flg desc
    --order by personi
    offset 0 row fetch next 1 rows only
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def indata(doc):
    etl = []
    etl.append({
        'personi':doc[0],
        'ID':doc[1],
        'Name':doc[2],
        'M':doc[3],
        'item':doc[4],        
        'occupation':doc[5],
        'level':doc[6],
        'license_no':doc[7],
        'effective_date':doc[8],
        'remark':doc[9],
        'status':doc[10],
        'updatetime':doc[11],
    })
    return etl

def toSQL(docs, totb, server, database, username, password):
    import pyodbc
    # conn_cmd = 'DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password
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
            
def capcha_resp(url,captchaImg,imgp,q,name,id,birdte_ad,birdte_ad_tw):
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
        #print(res)
        # post API then get result
        data = {
            'Form.NAME': f'{name}',
            'Form.IDNO': f'{id}',
            'Form.BIRDTE_AD': f'{birdte_ad}',
            'Form.BIRDTE_AD_TW': f'{birdte_ad_tw}',
            'Form.PNO': '',
            'Form.CRDITNO':'',
            'Form.ValidateCode': f'{res}',
        }
        
        resp = session.post(url,data=data)
        soup = BeautifulSoup(resp.text,"lxml")
        
        # 判斷辨識碼是否正確辨別
        result = str(soup.find('script', type='text/javascript'))
        
        if '驗證碼輸入錯誤' in str(result):
            session.close()
            q =+ 1
        else:
            session.close()
            return soup
            q = 100
            
def exit_obs(server,username,password,database,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count (distinct personi)
    from [{totb}]
    where  updatetime > = convert(varchar(10),getdate(),111)
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]


def delete(server,username,password,database,totb,id,personi):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    delete from [{totb}]
    where id = '{id}' and personi = '{personi}'
    
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()