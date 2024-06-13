# -*- coding: utf-8 -*-
"""
Created on Mon Sep 12 17:00:33 2022

@author: admin
"""
import os
server,database,username,password = 'vnt07.ucs.com','FRDB','pyuser','Ucredit7607'
def file_upload(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select distinct up.path 
    from treasure.skiptrace.dbo.casetb c(nolock),
    treasure.skiptrace.dbo.ClientTb ct(nolock),
    treasure.skiptrace.dbo.uploadimage up ,
    treasure.skiptrace.dbo.doctb tb 
    where c.clienti = ct.clienti and c.statusi not in (2,3,4)  and 
    c.casei = tb.casei and 
    tb.UploadImageI = up.UploadimageI and 
    ct.MasterClientI in (1000008,1000054,1000082,1000126,1000142) and 
    up.lastupdate>='2022/12/05' and 
    up.lastupdate<='2022/12/12'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

file = file_upload(server,username,password,database)
print(file)
#%%
file_exist = []
for i in file:


    filepath = i[0].lower().replace(rf'\\fortune\cashfile',rf'\\fortune\cashfile\client') 


    if os.path.isfile(filepath):
        file_exist.append("檔案存在。")
        flag = 'F'
    else:
        file_exist.append("檔案不存在。"+'_'+str(i[0])+'_'+filepath)
        flag = 'E'
        print(file_exist)
        
#%%
print(file_exist)

