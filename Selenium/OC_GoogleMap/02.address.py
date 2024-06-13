# -*- coding: utf-8 -*-
"""
Created on Tue Jan 31 14:48:54 2023

@author: admin
"""
import requests
from bs4 import BeautifulSoup
import time
import random
#sql srever db connection
db = {
    'server': 'vnt07.ucs.com',
    'database': 'UIS',
    'username': 'pyuser',
    'password': 'Ucredit7607',
    'fromtb':f'treasure.skiptrace.dbo.SVTb_PrioritySV_LIST',
    'totb':f'location_court',
}
#setting var
server,database,username,password,fromtb,totb = db['server'],db['database'],db['username'],db['password'],db['fromtb'],db['totb']

def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num
        
#經緯度須補數量
def src_obs(server,username,password,database,fromtb,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*) from {fromtb} where 經度 is null and 緯度 is null and 地址 not in (select 地址 from {totb}) and 地址 not like '%地號%' and 動機 not in (select 動機 from location_court where 動機 <> '' ) """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return obs[0][0]
#來源資料
def dbfrom(server,username,password,database,fromtb,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select * from {fromtb} where 經度 is null and 緯度 is null and 地址 not in (select 地址 from {totb}) and 地址 not like '%地號%' and 動機 not in (select 動機 from location_court where 動機 <> '' ) """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

#新增資料
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
#location_court
def location_court(doc):
    location_court = []
    location_court.append({
        "地址":doc[0],
        "緯度":doc[1],
        "經度":doc[2],
        "動機":doc[3],
    })
    return location_court
#更新資料
def updateOC(server,username,password,database,fromtb,address):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database,autocommit=True)
    cursor = conn.cursor()
    
    script = f"""
    update {fromtb}
    set [經度] = '' ,[緯度] = ''
    where [地址] = '{address}'
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()

obs = src_obs(server,username,password,database,fromtb,totb)
try:
    for i in foo(-1,obs-1):
        src = dbfrom(server,username,password,database,fromtb,totb)[0]
        address = src[11]
        address_tmp = src[11].replace('[','').replace(']','').replace('#','').replace('@','').replace('［','').replace('］','').replace('＃','').replace('＠','')
        print(address)
        motivation = src[17]
        print(motivation)
        #3+3
        html = requests.get('https://twzipcode.com',"html.parser")
        soup = BeautifulSoup(html.text)
        VIEWSTATEGENERATOR = soup.find("input",{"id":"__VIEWSTATEGENERATOR"}).get('value')
        EVENTVALIDATION = soup.find("input",{"id":"__EVENTVALIDATION"}).get('value')
        VIEWSTATE = soup.find("input",{"id":"__VIEWSTATE"}).get('value')
        payload  = {
        'Search_C_T': address_tmp,
        'submit': '找找',
        '__VIEWSTATEGENERATOR': VIEWSTATEGENERATOR ,
        '__EVENTVALIDATION': EVENTVALIDATION ,
        '__VIEWSTATE': VIEWSTATE }
        html1 = requests.post('https://twzipcode.com',data = payload)
        s = BeautifulSoup(html1.text)
        s1 = s.find_all("table")
        try:
            s2 = s1[1].find_all('td')
        except:
            updateOC(server,username,password,database,fromtb,address)
            continue
        try:
            try:
                #緯度
                if '緯度' in s2[14].get_text():
                    緯度 = s2[15].get_text()
                elif '緯度' in s2[16].get_text():
                    緯度 = s2[17].get_text()
                elif '緯度' in s2[18].get_text():
                    緯度 = s2[19].get_text()
                elif '緯度' in s2[20].get_text():
                    緯度 = s2[21].get_text()
                elif '緯度' in s2[22].get_text():
                    緯度 = s2[23].get_text()
                elif '緯度' in s2[24].get_text():
                    緯度 = s2[25].get_text()
                elif '緯度' in s2[26].get_text():
                    緯度 = s2[27].get_text()
                elif '緯度' in s2[28].get_text():
                    緯度 = s2[29].get_text()
                elif '緯度' in s2[30].get_text():
                    緯度 = s2[31].get_text()
                else :
                    緯度 = ''
            except:
                url='https://www.google.com/maps/place?q='+address_tmp
                #延遲時間可自行設定,但不可太過密集查詢,避免因影響他人使用而被封鎖
                time.sleep(random.randint(6,12))
                #執行資料抓取
                response = requests.get(url)
                soup = BeautifulSoup(response.text, "html.parser")
                text = soup.prettify() #text 包含了html的內容
                initial_pos = text.find(";window.APP_INITIALIZATION_STATE")
                data = text[initial_pos+36:initial_pos+85] #將其後的參數進行存取
                line = tuple(data.split(',')) #註1
                num1 = line[1] # latitude
                num2 = line[2] # longitude
                緯度  = num2
        except:
            updateOC(server,username,password,database,fromtb,address)
            continue
        try:
            try:
                #經度
                if '經度' in s2[14].get_text():
                    經度 = s2[15].get_text()
                elif '經度' in s2[16].get_text():
                    經度 = s2[17].get_text()
                elif '經度' in s2[18].get_text():
                    經度 = s2[19].get_text()
                elif '經度' in s2[20].get_text():
                    經度 = s2[21].get_text()
                elif '經度' in s2[22].get_text():
                    經度 = s2[23].get_text()
                elif '經度' in s2[24].get_text():
                    經度 = s2[25].get_text()
                elif '經度' in s2[26].get_text():
                    經度 = s2[27].get_text()
                elif '經度' in s2[28].get_text():
                    經度 = s2[29].get_text()
                elif '經度' in s2[30].get_text():
                    經度 = s2[31].get_text()
                else :
                    經度 = ''
            except:
                url='https://www.google.com/maps/place?q='+address_tmp
                #延遲時間可自行設定,但不可太過密集查詢,避免因影響他人使用而被封鎖
                time.sleep(random.randint(6,12))
                #執行資料抓取
                response = requests.get(url)
                soup = BeautifulSoup(response.text, "html.parser")
                text = soup.prettify() #text 包含了html的內容
                initial_pos = text.find(";window.APP_INITIALIZATION_STATE")
                data = text[initial_pos+36:initial_pos+85] #將其後的參數進行存取
                line = tuple(data.split(',')) #註1
                num1 = line[1] # latitude
                num2 = line[2] # longitude
                經度 = num1
        except:
           updateOC(server,username,password,database,fromtb,address)
           continue 
        
        #新增資料
        appdocs = (address,緯度,經度,motivation)
        print(appdocs)
        location_court_result = location_court(appdocs)
        toSQL(location_court_result, totb, server, database, username, password)
except:
    import smtplib
    from email.mime.text import MIMEText
    from email.header import Header
    sender = ['collection@ucs.com']
    receivers = ['DI@ucs.com'] # 接收郵件
    # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
    message = MIMEText('[ERROR] step 2 3+3 / google map 異常', 'plain', 'utf-8')
    message['From'] = Header('collection', 'utf-8') # 傳送者
    message['To'] = Header('DI', 'utf-8') # 接收者
    subject = '[ERROR] step 2 3+3 / google map 異常'
    message['Subject'] = Header(subject, 'utf-8')
    try:
        smtpObj = smtplib.SMTP('vrh19.ucs.com')
        smtpObj.sendmail(sender, receivers, message.as_string())
        print('郵件傳送成功')
    except smtplib.SMTPException:
        print('Error: 無法傳送郵件')

