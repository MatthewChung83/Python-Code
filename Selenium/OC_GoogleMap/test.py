# -*- coding: utf-8 -*-
"""
Created on Wed May 31 16:35:46 2023

@author: admin
"""


# -*- coding: utf-8 -*-
"""
Created on Wed May 31 09:58:44 2023

@author: admin
"""


import argparse
from PIL import Image
from os import *
import os
import shutil
import datetime




##Config
#外部參數輸入
parser = argparse.ArgumentParser()
parser.add_argument("arg1")
parser.add_argument("arg2")
args = parser.parse_args()
start_date = args.arg1
end_date = args.arg2
print(start_date)
print(end_date)
#DB
db = {
    'server': 'treasure',
    'database': 'skiptrace',
    'username': 'wuser',
    'password': 'wuser',
}

server,database,username,password= db['server'],db['database'],db['username'],db['password']


#select DB 資料
def dbfrom(server,username,password,database,start_date,end_date):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    Declare @st datetime = '{start_date}';
    Declare @et datetime = '{end_date}';
    ---------------------------
    -- 玉山戶謄反饋(CM)
    ---------------------------
    WITH Src as (
        select '聯合' as UCS,
                ct.ProductType as ProductTypeID,
                CASE WHEN ct.ClientName like '%新希望%' THEN '信放'
                    WHEN ct.ClientName like '%通信貸款%' THEN '信放'
                    WHEN ct.ClientName like '%NBA%' THEN '信放'
                    WHEN ct.ClientName like '%個人金融部%' THEN '信放'
                    ELSE '信卡' 
                END as ProductType,
                CASE WHEN ct.ClientName like '%新希望%' THEN '新希望'
                    WHEN ct.ClientName like '%通信貸款%' THEN '卡友貸'
                    WHEN ct.ClientName like '%NBA%' THEN '現金卡'
                    WHEN ct.ClientName like '%個人金融部%' THEN '信貸'
                    ELSE '信用卡'
                END as ProductTypeName,
                f.ID_NUMBER,
                f.NAME,
                f.CaseI,
                f.RetDate AS RetDateTime,
                d.DocI,
                CONVERT(VARCHAR(10), f.RetDate, 120) AS RetDate,
                d.DocDate,
                f.NO,
                f.STATUS,
                d.Content AS DocContent,
                p.PersonI,
                p.Name AS RelateName,
                p.Condition,
                ROW_NUMBER()OVER(Partition by p.ID Order by p.lastupdate desc) PersonSeq,
                CASE WHEN f.note != 'ESB' OR f.DPStatusNote != '線上FL' THEN 'N'
                     WHEN f.DPStatusNote = '線上FL' THEN 'Y'
                     ELSE ''
                END as IsOnlineApply,
    			d.UploadImageI
        from FLPOTb f
        JOIN ClientTb ct ON ct.clienti= f.client AND ct.MasterClientI = 1000028
        JOIN DocTb d on f.CaseI = d.CaseI
        JOIN (select * from PersonTb where ID is not null and ID != '') p on d.Content = p.ID
        WHERE d.Copy = '正本'
        AND f.RetStatus = 'Y'
        AND f.status = 'PULL FL'
        AND d.DocType = '戶籍謄本'
        AND f.retdate BETWEEN @st and @et
        AND d.DocDate BETWEEN @st and @et
    ),
    ResultTmp AS (
        SELECT distinct 
               UCS
             , ProductType
             , ProductTypeName
             , ID_NUMBER
             , NAME
             , CaseI
             , RetDate
             , DocI
             , DocDate
             , DocContent
             , IsOnlineApply
    		 , UploadImageI
        FROM Src
        WHERE Condition not in ('去世', '死亡')
        AND IsOnlineApply != ''
    )
    select distinct
     s3.Path as [SrcFilePath]
    	, '\\fortune\'+'UCS'+'\DI\test\' + CONVERT(varchar(8), @et, 112) + '\' + s1.ID_NUMBER + '.pdf' as [TargetFilePath]
    from ResultTmp s1
    left join (select * from USR_Data_REG where dCreateTime BETWEEN @st and @et) s2 on s1.DocContent = s2.cId1
    left join UploadImage s3 on s1.UploadImageI = s3.UploadimageI"""    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

try:
    d = datetime.datetime.strptime(end_date, "%Y/%m/%d")
except:
    d = datetime.datetime.strptime(end_date, "%Y-%m-%d")
s = d.strftime('%Y%m%d')
print(s)
os.makedirs(rf'\\fortune\\AAADbases\\Return\\ESUN\FL\\'+s)
print(s)


f = dbfrom(server,username,password,database,start_date,end_date)

for i in f :
    input_file = i[0]
    output_file = str('\\'+i[1])
    print(input_file,output_file)
    try:
        shutil.copyfile(input_file, output_file)
        

            
    except:
        picture = Image.open(input_file)
        picture.save(output_file, "PDF", save_all=True)
        pass



