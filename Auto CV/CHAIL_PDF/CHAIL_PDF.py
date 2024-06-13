# -*- coding: utf-8 -*-
"""
Created on Thu Oct  5 16:15:30 2023

@author: admin
"""


import pdfkit
import pandas as pd
import pymssql
import os
import argparse

parser = argparse.ArgumentParser(description = 'manual to this script')
parser.add_argument('--date',type = str,default = None)
args = parser.parse_args()
db = {
    'server': 'RICHES',
    'database': 'UCS_ETL',
    'username': 'CLUSER',
    'password': 'Ucredit7607',
}
Data_dt = args.date
Year = Data_dt.split('-')[0]
Date = Data_dt.split('-')[1]+Data_dt.split('-')[2]+'(帳單)'
Folder = '中租合迪'
try:
    os.makedirs(rf'\\fortune\Cashfile\{Year}\ReferImage\中租合迪')
except:
    pass
try:
    os.makedirs(rf'\\fortune\Cashfile\{Year}\ReferImage\中租合迪\{Date}')
except:
    pass

server,database,username,password= db['server'],db['database'],db['username'],db['password']
def data(server,username,password,database,Data_dt):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
        SELECT 
          DISTINCT [CaseI]
        FROM [UCS_ETL].[dbo].[ODS_Chailease_BillData]	
        WHERE DataDt = '{Data_dt}'
        """
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src
src = data(server,username,password,database,Data_dt)

for i in range(len(src)):
    ID = src[i][0]
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    script = f"""
        SELECT 
          [契編] = ContractNumber
          ,[繳款日期] = isnull(PayDate,'')
          ,[繳款金額] = isnull(PayAmount,'')
          ,[沖帳日期] = isnull(BalanceDate,'')
          ,[沖帳金額] = isnull(BalanceAmount,'')
          ,[繳款方式] = isnull(PayMethod,'')
        FROM [UCS_ETL].[dbo].[ODS_Chailease_BillData]	
        WHERE DataDt = '{Data_dt}' and CaseI = '{ID}'
        ORDER BY PayDate
        """
    df = pd.read_sql(script,conn)
    print(df)
    f = open('exp.html','w')
    a = df.to_html(index=False)
    f.write(a)
    f.close()
    config = pdfkit.configuration(wkhtmltopdf="C:\Program Files\wkhtmltopdf\wkhtmltopdf.exe")
    pdfkit_options = {'encoding':"big5"}
    pdfkit.from_file('exp.html', rf'\\fortune\Cashfile\{Year}\ReferImage\中租合迪\{Date}\{ID}.pdf',configuration=config,options= pdfkit_options)
    


    