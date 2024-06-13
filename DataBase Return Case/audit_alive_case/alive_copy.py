# -*- coding: utf-8 -*-
"""
Created on Thu Dec  1 17:48:59 2022

@author: admin
"""


# -*- coding: utf-8 -*-
"""
Created on Mon Feb 21 18:42:45 2022

@author: matthew5043
"""
import pyodbc
import pandas as pd

from datetime import datetime, date, timedelta

yesterday = date.today() + timedelta(days = -7)
today = date.today() + timedelta(days = 0)

client = ['1000008','1000054','1000082','1000126','1000142']
folder = ['CTB','TFIT','FEIB','DBS','X\HSBC']
server = 'vnt07.ucs.com' 
database = 'FRDB' 
username = 'pyuser' 
password = 'Ucredit7607'  
cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

def check_obs_T(server,username,password,database,yesterday):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""

    select count(distinct up.path) from treasure.skiptrace.dbo.casetb c(nolock),
    treasure.skiptrace.dbo.ClientTb ct(nolock),
    treasure.skiptrace.dbo.uploadimage up ,
    treasure.skiptrace.dbo.doctb tb where c.clienti = ct.clienti and 
    c.statusi not in (2,3,4)  and c.casei = tb.casei and tb.UploadImageI = up.UploadimageI 
    and ct.MasterClientI in (1000008,1000054,1000082,1000126,1000142) and up.lastupdate>='{str(yesterday)}'
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return list(obs[0])[0]
obs_T = check_obs_T(server, username, password, database,yesterday)
#print(obs_T)
def check_obs_Y(server,username,password,database,yesterday,today):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""

    select count(distinct up.path) from treasure.skiptrace.dbo.casetb c(nolock),
    treasure.skiptrace.dbo.ClientTb ct(nolock),
    treasure.skiptrace.dbo.uploadimage up ,
    treasure.skiptrace.dbo.doctb tb where c.clienti = ct.clienti and 
    c.statusi not in (2,3,4)  and c.casei = tb.casei and tb.UploadImageI = up.UploadimageI 
    and ct.MasterClientI in (1000008,1000054,1000082,1000126,1000142) and up.lastupdate>='{str(yesterday)}' 
    
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return list(obs[0])[0]
obs_Y = check_obs_Y(server, username, password, database,yesterday,today)
#print(obs_Y)

for x in range(len(client)):
    a = client[x]
    #b=src_obs(server, username, password, database, a)
    #df = pd.read_sql(b, cnxn)
    #print(df)
    import pyodbc
    import pandas as pd

    server = 'vnt07.ucs.com' 
    database = 'UIS'
    username = 'pyuser' 
    password = 'Ucredit7607'  
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
    cursor = cnxn.cursor()
    query = "select distinct up.path from treasure.skiptrace.dbo.casetb c(nolock),treasure.skiptrace.dbo.ClientTb ct(nolock),treasure.skiptrace.dbo.uploadimage up ,treasure.skiptrace.dbo.doctb tb where c.clienti = ct.clienti and c.statusi not in (2,3,4)  and c.casei = tb.casei and tb.UploadImageI = up.UploadimageI and ct.MasterClientI =" +str(a)+" and up.lastupdate>=" +"'"+str(yesterday)+"'" 
    
    df = pd.read_sql(query, cnxn)

    import os
    import shutil

    for i in df.index: 
        try:
            file = df.loc[i,'path'].lower()
                   
            #print(rf'\\fortune\CashFile\client\\'+str(folder[x]))
    
            file1 = file.replace(rf'\\fortune\cashfile',rf'\\fortune\CashFile\client\\'+str(folder[x]))
            
            ans = os.path.split(file1)[0]
            #print(file1)
            try:
                os.makedirs(ans)
            except:
                pass
            shutil.copy2(file, file1)
        except:
            if os.path.isfile(file):
                print("來擋檔案存在。",file)
            else:
                print("來擋檔案不存在。",file)
            if os.path.isfile(file1):
                print("目的檔案存在。",file1)
            else:
                print("目的檔案不存在。",file1)
            #print('copy faild',i)
            continue


