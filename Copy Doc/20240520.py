# -*- coding: gb2312 -*-
"""
Created on Wed Nov 30 16:07:36 2022

@author: admin
"""
import pyodbc
import pandas as pd

server = 'RICHES' 
database = 'CL_Daily'
username = 'CLUSER' 
password = 'Ucredit7607'  
cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
query = f"""
 select distinct s2.ID,s2.ID+'\\'+REVERSE(dbo.SplitX(REVERSE(s2.Path),'\\',0))as new_doc,convert(varchar(4000),s2.Path)as Path,doci from treasure.skiptrace.dbo._DocTb_1000016 s1
    left join treasure.skiptrace.dbo.UploadImage s2 on s1.UploadImageI = s2.UploadimageI
    where path is not null
"""


df = pd.read_sql(query, cnxn)
import os
import shutil

for i in df.index: 
    file = df.loc[i,'Path'].lower()
    new_file = file.replace('\\fortune\\cashfile','\\collection\cf$')
    New_doc = df.loc[i,'new_doc'].lower()
    ID = df.loc[i,'ID']
    doci = df.loc[i,'doci']
    ta_file = rf'C:\Users\matthew5043.UCS\Desktop\SQL Script\20240520\{str(ID)}'
    ta_file1 = rf'C:\Users\matthew5043.UCS\Desktop\SQL Script\20240520\{str(New_doc)}'

    ans = os.path.split(ta_file1)[0]
    

    try:
        os.mkdir(ta_file)
    except:
        pass
    try:
        try:
            shutil.copy(file,ta_file1)
        except:
            shutil.copy(new_file,ta_file1)
    except:
        log = 'Not Found '+str(doci)+'\n'
        path = rf'C:\Users\matthew5043.UCS\Desktop\SQL Script\20240520\Error_FILE.txt'
        f = open(path, 'a')
        try:
            f.writelines(log)
        except:
            pass
