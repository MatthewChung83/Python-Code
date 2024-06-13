from operator import itemgetter, attrgetter
from datetime import datetime, timedelta
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
import pandas as pd
import requests
import time
import pdfkit

from etl_func import *
from config import *
from etl_func import *
from dict import *

server,database,username,password,totb1,totb2 = db['server'],db['database'],db['username'],db['password'],db['totb1'],db['totb2']
url = wbinfo['url']

# 分關鍵字查詢
try:
    for i in range(len(qtypes)):
        qtype = qtypes[i]
        
        data = {
            'Action': 'Qeury',
            'lstoken': '9996409954',
            'Q_DMDeptMainID':'',
            'TBOXDMPostDateS': str_date,
            'TBOXDMPostDateE': end_date,
            'Q_DTACaseNumber':'',
            'Q_DMTitle': '',
            'Q_DTACatCode1':'',
            'Q_DTAColumn4':'',
            'Q_DMCatCode':'',
            'Q_DMBody': qtype,
            'BtnSubmit': '查詢',
            }
        
        resp = requests.post(url,data=data)
        soup = BeautifulSoup(resp.text,"lxml")
        
        if soup.find(class_ = 'page'):
            
            # 取得分頁列表
            res_page = soup.find(class_ = 'page').find_all('a', href=True)
            page_links = []
            for page in res_page:
                page_links.append(page['href'])
            page_links = list(set(page_links))
            
            # 分頁爬取
            for page_link in page_links:
                resp_page = requests.get('https://www.judicial.gov.tw/tw/' + str(page_link))
                #print(resp_page.text)
                soup_page = BeautifulSoup(resp_page.text,"lxml")
                judicials = soup_page.find('tbody').find_all('tr')
                
                # 分案爬取
                for judicial in judicials:
                    resp_doc = requests.get('https://www.judicial.gov.tw/' + str(judicial.find_all('a' , href = True)[0]['href']))
                    print('https://www.judicial.gov.tw/' + str(judicial.find_all('a' , href = True)[0]['href']))
                    soup_doc = BeautifulSoup(resp_doc.text,"lxml")
                    judicial_block = soup_doc.find(class_ = 'content_block')
                    
                    # 定義頁面欄位
                    court = judicial.find_all('td')[1].text.strip()
                    judicial_no = judicial.find_all('td')[2].text.strip()
                    recipient = judicial.find_all('td')[3].text.strip()
                    area = judicial.find_all('td')[4].text.strip()
                    doc_type = judicial.find_all('td')[5].text.strip()
                    publishing_date = judicial.find_all('td')[6].text.strip()
                    judicial_type = judicial.find_all('td')[7].text.strip()
    
                    # 判決文件解析
                    judicial_title = judicial_block.find('h2').text.strip()
                    judicial_content = judicial_block.find(class_ = "cp").text.strip()
                    _judicial_content = judicial_content.replace(' ','').replace('　','').replace(' ','')
                    
                    # 重置分詞順序
                    paragraphs,seq = [],1
                    for word in segwords:
                        paragraphs.append([seq,word,_judicial_content.find(word)])
                        seq += 1
                    paragraphs = sorted(paragraphs, key = itemgetter(2))
    
                    # 分詞斷句
                    for i in range(len(paragraphs)):
                        if paragraphs[i][2] < 0:
                            paragraphs[i].append('')
                        else:
                            if i == len(paragraphs) - 1:
                                paragraphs[i].append(_judicial_content[_judicial_content.find(paragraphs[i][1]):10000])
                            else:
                                paragraphs[i].append(_judicial_content[_judicial_content.find(paragraphs[i][1]):_judicial_content.find(paragraphs[i+1][1])])
                    paragraphs = sorted(paragraphs, key = itemgetter(0))
                    
                    # 定義文件內容欄位
                    if paragraphs[0][3].strip() == '' or paragraphs[0][3].strip() == ' ' or paragraphs[0][3].strip() is None:
                        judicial_doc_date = publishing_date
                    else:
                        judicial_doc_date = paragraphs[0][3].strip()
                        
                    if paragraphs[1][3].strip() == '' or paragraphs[1][3].strip() == ' ' or paragraphs[1][3].strip() is None:
                        judicial_doc_no = judicial_no
                    else:
                        judicial_doc_no = paragraphs[1][3].strip()
                        
                    judicial_doc_subject = paragraphs[2][3].strip()
                    judicial_doc_basis = paragraphs[3][3].strip()
                    judicial_doc_anncm = paragraphs[4][3].strip()
                    judicial_doc_pdate = paragraphs[5][3].strip()
                    judicial_doc_udate = paragraphs[6][3].strip()
                    judicial_doc_unit = paragraphs[7][3].strip()
                                 
                    # 債務人/被告/繼承人等人員清單                
                    names = Parse_Title(i,judicial_title,splits) + ',' + Parse_Subject(i,judicial_doc_subject,splits) + ',' + Parse_anncm(i,judicial_doc_anncm,splits)
                    result = delws_re(names,delws,delimis)
                    
                    html = 'https://www.judicial.gov.tw/' + str(judicial.find_all('a' , href = True)[0]['href'])
                    config = pdfkit.configuration(wkhtmltopdf="C:\Program Files\wkhtmltopdf\wkhtmltopdf.exe")
                    file_save =  rf'C:\Py_Project\project\Judicial\file\\'+court+'_'+judicial_no+'.pdf'
                    pdfkit.from_url(html,file_save,configuration=config)
                    
                    judicial_content = judicial_content.replace('\r\n','').replace('\xa0','').replace('├','').replace('─','').replace('┤','').replace('┼','').replace('│','').replace('└','').replace('┴','').replace('┘','')
    
                    # 入資料庫
                    if src_obs(server,username,password,database,judicial_no,judicial_doc_no,publishing_date) == 0:
                        
                        insertdate = (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
                        
                        detail = [judicial_no,recipient,publishing_date,judicial_doc_date,judicial_doc_no,judicial_doc_subject,result[0],result[1],insertdate]
                        doc2 = indetail(detail)
                        
                        data = [qtype,court,judicial_no,recipient,area,doc_type,publishing_date,judicial_type,judicial_title,judicial_doc_date,judicial_doc_no,judicial_doc_subject,
                                judicial_doc_basis,judicial_doc_anncm,judicial_doc_pdate,judicial_doc_udate,judicial_doc_unit,judicial_content,insertdate,result[0],file_save]
                        doc1 = indata(data)
    
                        # 標題檔
                        toSQL(doc1, totb1, server, database, username, password)
                        
                        # # 債務人檔
                        toSQL(doc2, totb2, server, database, username, password)
except:
    pass
import shutil
import os
import time
import datetime


file_source = rf'C:\\Py_Project\\project\\Judicial\\file\\'
file_destination = rf'\\fortune\Cashfile\UCS\Judicial\\'
 
get_files = os.listdir(file_source)

time.strftime('%Y%m%d')
now_time = datetime.datetime.now()
future_time = now_time + datetime.timedelta(days=-4)
fu = future_time.strftime('%Y%m%d')
print(fu)

for g in get_files:
    t = time.localtime(os.path.getctime(file_source+g))
    ctime=time.strftime("%Y%m%d",t)
    
    if ctime > fu :     
        shutil.copy(file_source + g, file_destination)
        print(ctime)
        print(g)
        pass
    
    else :

        pass
#Send Mail        
import smtplib
from email.mime.text import MIMEText
from email.header import Header
sender = 'matthew5043@ucs.com'
receivers = ['matthew5043@ucs.com'] # 接收郵件
# 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
message = MIMEText('新資源Judicial上載圖檔已完成', 'plain', 'utf-8')
message['From'] = Header('wbt', 'utf-8') # 傳送者
message['To'] = Header('matthew5043', 'utf-8') # 接收者
subject = '新資源Judicial上載圖檔已完成'
message['Subject'] = Header(subject, 'utf-8')
try:
    smtpObj = smtplib.SMTP('vrh19.ucs.com')
    smtpObj.sendmail(sender, receivers, message.as_string())
    print('郵件傳送成功')
except smtplib.SMTPException:
    print('Error: 無法傳送郵件')