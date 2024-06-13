# -*- coding: utf-8 -*-
"""
Created on Wed Jul 19 14:33:25 2023

@author: admin
"""
import re
from pdfminer.high_level import extract_text

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image


db = {
    'server': 'RICH',
    'database': 'skiptrace',
    'username': 'FRUSER',
    'password': '1qaz@WSX',
    'database1': 'UCS_ReportDB',
    'username1': 'FRUSER',
    'password1': '1qaz@WSX',
    'totb' : 'INS_Consumer_DTL',
}

server,database,username,password= db['server'],db['database'],db['username'],db['password']
database1,username1,password1,totb = db['database1'],db['username1'],db['password1'],db['totb']

def data(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select distinct
        ct.MasterClientI,
        c.CaseI,
        up.ID,
        s.Description,
        e.name,
        isnull(c.BankExecuteGroundStatus,'N'),
        case when exists(select * from LDCTb (nolock)where casei = c.casei and DebtorID = up.ID and Code = 'RD' and RequestDate > =(select dateadd(m, datediff(m,0,getdate())-12,0)) ) then 'Y' else 'N' end,
        case when exists(select * from LDCTb (nolock)where casei = c.casei and DebtorID = up.ID and Code = 'AA'and RequestDate > =(select dateadd(m, datediff(m,0,getdate())-12,0)) ) then 'Y' else 'N' end,
        '' as [YN],
        getdate(),
        STUFF((SELECT ',' + Path
                FROM UploadImage(nolock)
                WHERE CaseI = d.CaseI and Doctype = '消費明細表' order by CaseI
                FOR XML PATH('')), 1, 1, '')
    from casetb c (nolock)
    left join clienttb ct (nolock)on c.clienti = ct.clienti 
    left join doctb d (nolock)on c.CaseI = d.CaseI
    left join UploadImage up (nolock)on d.UploadImageI = up.UploadimageI
    left join statustb s (nolock)on c.StatusI = s.StatusI
    left join emptb e (nolock)on c.empi = e.empi
    where 
        ct.MasterClientI = 1000082 and c.statusi not in (2,3,4) and d.doctype ='消費明細表' and up.ID is not null and up.ID not in (select distinct  ID from UCS_ReportDB.dbo.INS_Consumer_DTL where Date> = getdate())
        and d.docdate> = (select dateadd(m, datediff(m,0,getdate())-1,0))



	
	
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def read_pdf_file(file_path):
    text = extract_text(file_path)
    return text


def INS(doc):
    INS= []

    INS.append({
        "MasterClientI":doc[0],
        "Casei":doc[1],
        "ID":doc[2],
        "STATUS":doc[3],
        "Emp_Name":doc[4],
        "Bank_YN":doc[5],
        "RD_YN":doc[6],
        "AA_YN":doc[7],
        "CON_YN":doc[8],
        "date":doc[9],
    })
    return INS
def toSQL(docs,totb,server, database, username, password):
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
def convert_jpg_to_pdf(image_path, pdf_path):
    # 打開圖片並獲取其大小
    image = Image.open(image_path)
    width, height = image.size
    # 創建 PDF 文件
    c = canvas.Canvas(pdf_path, pagesize=letter)
    c.drawImage(image_path, 0, 0, width, height)  # 你也可以改變圖片在 PDF 中的大小和位置
    c.save()
    
df =  data(server,username,password,database)

for i in range(len(df)):
    retry = 1
    while retry < 10 :
        try:
            MasterClientI = df[i][0]
            Casei = df[i][1]
            ID = df[i][2]
            STATUS = df[i][3]
            Emp_Name = df[i][4]
            BankStatus = df[i][5]
            RD = df[i][6]
            AA = df[i][7]
            YN = 'N'
            date = df[i][9]
            Path = df[i][10]
            for i in Path.split(','):
                if '.pdf' not in i :
                     # 轉換圖片並保存結果
                    convert_jpg_to_pdf(i, 'output.pdf')
                    pdf_content = read_pdf_file('output.pdf')
                    keywords = re.findall(r'人壽|保險|保單|保費', pdf_content)
                    print(i)
                else:
                    pdf_content = read_pdf_file(i)
                    keywords = re.findall(r'人壽|保險|保單|保費', pdf_content)
                    print(i)
                if len(keywords) > 0:
                    YN = 'Y'
                    continue
        
            docs = (MasterClientI,Casei,ID,STATUS,Emp_Name,BankStatus,RD,AA,YN,date)
            INS_result = INS(docs)
            toSQL(INS_result, totb, server, database1, username1, password1)
            retry = 10
        except:
            retry = retry + 1
            pass
