def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num

def src_obs(server,username,password,database,totb1,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*)
    from [dbo].[{totb1}]
    where entitytype in ('{entitytype}') and (status is null)
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]

def dbfrom(server,username,password,database,totb1,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select *
    from [dbo].[{totb1}] as a
    where entitytype in ('{entitytype}') and (status is null)
    order by rowi
    offset 0 row fetch next 1 rows only
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def updateSQL(server,username,password,database,totb1,entitytype,status,updatetime,ID,driver_type,driver_status,DRvaliddate):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    update [dbo].[{totb1}]
    set driver_type = '{driver_type}', driver_status = '{driver_status}', DRvaliddate = '{DRvaliddate}',status = '{status}', updatetime = '{updatetime}'
    where entitytype in ('{entitytype}') and ID = '{ID}'
    """
    
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()

def capcha_resp(url,captchaImg,imgp,q,ID,birthday):
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
            'stage': 'natural',
            'method': 'queryResult',
            'uid': f'{ID}',
            'birthday': f'{birthday}',
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