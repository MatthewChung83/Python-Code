# -*- coding: utf-8 -*-
"""
Created on Tue Jul 25 10:20:32 2023

@author: admin
"""

import os
import shutil


from PyPDF2 import PdfFileReader 
from PyPDF2 import PdfFileWriter 


from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image


db = {
    'server': '10.10.0.94',
    'database': 'skiptrace',
    'username': 'CLUSER',
    'password': 'Ucredit7607',
    
}
server,database,username,password = db['server'],db['database'],db['username'],db['password']

def dbfrom(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
        select
        	c.AcctI_1,
        	(convert(nvarchar(Max),up.Path)),
        	c.casei,
        	d.DocDate
        from casetb c,UploadImage up,doctb d
        where c.CaseI = up.CaseI and up.UploadimageI= d.UploadImageI and up.CaseI = d.CaseI and
        	c.clienti in (select clienti from clienttb where MainClientI in (1002217))
        	and c.StatusI not in (2,3,4) and c.casei = 11851211--d.docdate > = convert(varchar(10),getdate()-1,111)
            
        
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def dbfrom_del(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
        select distinct 
            c.AcctI_1,
            (convert(nvarchar(Max),up.Path)),
            c.casei,
            d.DocDate
        from casetb c,UploadImage up,doctb d
        where c.CaseI = up.CaseI and up.UploadimageI= d.UploadImageI and up.CaseI = d.CaseI and
            c.clienti in (select clienti from clienttb where MainClientI in (1002217))
            and c.StatusI not in (2,3,4)
            and datediff(dd,d.DocDate,getdate())> = 30
            
        
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def dbfrom_close(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
        select
            c.AcctI_1
            
        from casetb c
        where 
            c.clienti in (select clienti from clienttb where MainClientI in (1002217))
            and c.StatusI in (0,2,3,4,19,59,141)

            
        
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def convert_jpg_to_pdf(image_path, pdf_path):
    # 打開圖片並獲取其大小
    image = Image.open(image_path)
    width, height = image.size
    # 創建 PDF 文件
    c = canvas.Canvas(pdf_path, pagesize=letter)
    c.drawImage(image_path, 0, 0, width, height)  # 你也可以改變圖片在 PDF 中的大小和位置
    c.save()
    
def resize_image(image_path, output_path, max_size):
    # 打開圖像
    image = Image.open(image_path)
    # 改變圖像的大小
    image.thumbnail(max_size)
    # 儲存結果
    image.save(output_path)



data = dbfrom(server,username,password,database)
for i in range(len(data)):
    AcctI_1 = data[i][0]
    casei = data[i][2]
    path = data[i][1].lower()
    path_c = path.replace(str(casei),str(AcctI_1))
    LastUpdate = data[i][3]
    print(path)
    copy_path = path_c.replace(rf'\\fortune\cashfile',rf'\\fortune\SCB\4-圖檔\{AcctI_1}')
    mkdir_path = copy_path.replace(copy_path.rsplit('\\', 1)[1],'')
    
    copy_path_tiff = path_c.replace(rf'\\fortune\cashfile',rf'\\fortune\SCB\4-圖檔\{AcctI_1}').replace('tif','pdf')
    copy_path_jpg = path_c.replace(rf'\\fortune\cashfile',rf'\\fortune\SCB\4-圖檔\{AcctI_1}').replace('jpg','pdf')
    
    #建資料夾
    try :
        os.makedirs(mkdir_path, exist_ok=True)
    except:
        pass
    
    
    #try:
    if '.pdf' not in path :
        print(path)
        if '.tif' in path:
            # 改變圖像的大小並保存結果
            resize_image(path, 'resized_image.tiff', (800, 600))
            # 轉存PDF
            convert_jpg_to_pdf('resized_image.tiff', copy_path_tiff)
            # 加密6020
            pdf_reader = PdfFileReader(copy_path_tiff)
            pdf_writer = PdfFileWriter()
            for page in range(pdf_reader.getNumPages()):
                pdf_writer.addPage(pdf_reader.getPage(page))
            pdf_writer.encrypt('6020')
            with open(copy_path_tiff,'wb') as out:
                pdf_writer.write(out)
                
        elif '.jpg' in path:
            # 改變圖像的大小並保存結果
            resize_image(path, 'resized_image.jpg', (800, 600))
            # 轉存PDF
            convert_jpg_to_pdf('resized_image.jpg', copy_path_jpg)
            # 加密6020
            pdf_reader = PdfFileReader(copy_path_jpg)
            pdf_writer = PdfFileWriter()
            for page in range(pdf_reader.getNumPages()):
                pdf_writer.addPage(pdf_reader.getPage(page))
            pdf_writer.encrypt('6020')
            with open(copy_path_jpg,'wb') as out:
                pdf_writer.write(out)
         
        
    else:
        # 加密6020
        pdf_reader = PdfFileReader(path)
        pdf_writer = PdfFileWriter()
        for page in range(pdf_reader.getNumPages()):
            pdf_writer.addPage(pdf_reader.getPage(page))
        pdf_writer.encrypt('6020')
        with open(copy_path,'wb') as out:
            pdf_writer.write(out)
    #except:
        #pass
#超過30天刪除檔案
data_del = dbfrom_del(server,username,password,database)
for i in range(len(data_del)):
    AcctI_1_del = data_del[i][0]
    casei_del = data_del[i][2]
    path_del = data_del[i][1].lower()
    path_c_del = path_del.replace(str(casei_del),str(AcctI_1_del))
    copy_path_del = path_c_del.replace(rf'\\fortune\cashfile',rf'\\fortune\SCB\4-圖檔\{AcctI_1_del}').replace('.jpg','.pdf').replace('.tif','.pdf')
    print(copy_path_del)    
    # 
    try:
        os.remove(copy_path_del)
    except:
        pass

#關檔案件即刪除資料夾
data_close = dbfrom_close(server,username,password,database)

for i in range(len(data_close)):
    AcctI_1_close = data_close[i][0]
    print(AcctI_1_close)
    # 指定要刪除的目錄
    dir_path = rf'\\fortune\SCB\4-圖檔\{AcctI_1_close}'
    
    # 如果目錄存在，則刪除它
    if os.path.exists(dir_path):
        shutil.rmtree(dir_path)


# 指定要檢查的目錄
dir_path = rf'\\fortune\SCB\4-圖檔'

# 遍歷目錄樹
for root, dirs, files in os.walk(dir_path, topdown=False):
    # 如果一個目錄為空，則刪除它
    for name in dirs:
        full_path = os.path.join(root, name)
        if not os.listdir(full_path):
            os.rmdir(full_path)
















