# -*- coding: utf-8 -*-
"""
Created on Tue Dec 28 14:56:44 2021

@author: matthew5043
"""
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice
from pdfminer.layout import *
from pdfminer.converter import PDFPageAggregator
import pandas as pd
#from collections.abc import Iterable
#import warnings
#warnings.simplefilter('ignore', DeprecationWarning)
import pymssql
import pandas as pd
from os import listdir
import csv
from os.path import isfile, isdir, join
import datetime
from datetime import timedelta
now = datetime.datetime.now()
this_week_start = now - timedelta(days=now.weekday())
FD=this_week_start.strftime('%Y%m%d')
weekly='AM_Weekly_89850389_'+FD+'_EL'
#weekly='AM_Weekly_89850389_'+'20230619'+'_EL'
print(weekly)
#weekly = 'AM_Weekly_89850389_20220502_EL'

db_server = 'vnt07'
db_acct = 'pyuser'
db_pwd = 'Ucredit7607'

# 取包含'EL'字樣的檔名
doc = 'AM'
# 取副檔名為pdf的檔名
comnm = '.pdf'
# 取pdf檔案存放之目錄

mypath = rf'C:\Py_Project\project\easymap_EL_1.2\gen_el05\file\{weekly}'
#mypath = r'\\vnt07\ucsfile\el\AM_Weekly_89850389_20201207_EL'
# 文件類型spec (後續須依照PDF內容擴增)
spec_ref = r'C:\Py_Project\env\el_spec.csv'
# 例外內容每次執行時之log檔(可利用例外內容去擴充el_spec.csv內容)(每次執行時會抓去無法辨識的新欄位)
exception = rf'{mypath}\exception.csv'

path = f"{mypath}\\"

def findchar(subchar,chars,n):
    count = 0
    while n > 0:
        index = chars.find(subchar)
        if index <= 0:
            return -1
        else:
            chars = chars[index+1:]
            n -= 1
            count = count + index + 1
    return count - 1

def AWCsv(pathfile,data):
    csvFileToWrite = open(pathfile, 'a', encoding='utf-8-sig', newline='')
    csvDataToWrite = csv.writer(csvFileToWrite)
    csvDataToWrite.writerow(data)
    csvFileToWrite.close()

def docspec(pathfile,item,type_desc):
    allspec = pd.read_csv(pathfile).drop_duplicates()
    if item == '':
        spec = allspec
    else :
        itemfilter = allspec['item'] == item
        spec = allspec[itemfilter]
    spec_len = len(spec)
    return spec

# 寫入資料庫(土地標示部)，參數:table,欄位1,欄位2,...
def Insertlandlabel(pdf_nm,title,lot,print_dttm,land_item,type_seq,reg_date,reg_reason,area,use_partition,type_of_use,present_value,building_no,other_reg_items):
    conn = pymssql.connect(server=db_server, user=db_acct, password=db_pwd)
    cursor2 = conn.cursor()
    script = f"""
    if not exists(select * from [UIS].[dbo].[Elpdf_{land_item}] where pdf_nm = '{pdf_nm}' and print_dttm = '{print_dttm}' and type_seq = {type_seq})
        insert into [UIS].[dbo].[Elpdf_{land_item}]
        (pdf_nm,title,lot,print_dttm,land_item,type_seq,reg_date,reg_reason,area,use_partition,type_of_use,present_value,building_no,other_reg_items)
        values ('{pdf_nm}','{title}','{lot}','{print_dttm}','{land_item}',{type_seq},'{reg_date}','{reg_reason}','{area}','{use_partition}','{type_of_use}','{present_value}','{building_no}','{other_reg_items}')
    """
    cursor2.execute(script)
    conn.commit()
    cursor2.close()
    conn.close()

# 寫入資料庫(土地所有權部)，參數:table,欄位1,欄位2,...
def Insertlandown(pdf_nm,title,lot,print_dttm,land_item,type_seq,reg_order,reg_date,reg_reason,cause_date,owner,id,rights_scope,title_no,declared_price,pretransfer_value_or_original_price,rights_obtained_scope,reg_other_of_other_right,other_reg_items):
    conn = pymssql.connect(server=db_server, user=db_acct, password=db_pwd)
    cursor2 = conn.cursor()
    script = f"""
    if not exists(select * from [UIS].[dbo].[Elpdf_{land_item}] where pdf_nm = '{pdf_nm}' and print_dttm = '{print_dttm}' and type_seq ={type_seq})
        insert into [UIS].[dbo].[Elpdf_{land_item}]
        (pdf_nm,title,lot,print_dttm,land_item,type_seq,reg_order,reg_date,reg_reason,cause_date,owner,id,rights_scope,title_no,declared_price,pretransfer_value_or_original_price,rights_obtained_scope,reg_other_of_other_right,other_reg_items)
        values ('{pdf_nm}','{title}','{lot}','{print_dttm}','{land_item}',{type_seq},'{reg_order}','{reg_date}','{reg_reason}','{cause_date}','{owner}','{id}','{rights_scope}','{title_no}','{declared_price}','{pretransfer_value_or_original_price}','{rights_obtained_scope}','{reg_other_of_other_right}','{other_reg_items}')
    """
    cursor2.execute(script)
    conn.commit()
    cursor2.close()
    conn.close()

# 寫入資料庫(土地他項權利部)，參數:table,欄位1,欄位2,...
def Insertlandother(pdf_nm,title,lot,print_dttm,land_item,type_seq,reg_order,delayed_interest,co_guarantee_area_no,co_guarantee_building_no,rent,paper_no,duration,receipted_year,address,interest_rate,instructions,other_guarantee_scope,property_by_payoff,settlement_date,id,set_purpose,reg_date,reg_reason,debt_ratio,damaged_fee,situation_of_prepaid,subject_reg_order,secured_types_and_scope,secured_confirmation_date,ttl_amount_of_secured,certificate_no,right_person,rights_type,rights_value,rights_subject,rights_scope,assign_setting_restrictions,other_reg_items):
    conn = pymssql.connect(server=db_server, user=db_acct, password=db_pwd)
    cursor2 = conn.cursor()
    script = f"""
    if not exists(select * from [UIS].[dbo].[Elpdf_{land_item}] where pdf_nm = '{pdf_nm}' and print_dttm = '{print_dttm}' and type_seq ={type_seq})
        insert into [UIS].[dbo].[Elpdf_{land_item}]
        (pdf_nm,title,lot,print_dttm,land_item,type_seq,reg_order,delayed_interest,co_guarantee_area_no,co_guarantee_building_no,rent,paper_no,duration,receipted_year,address,interest_rate,instructions,other_guarantee_scope,property_by_payoff,settlement_date,id,set_purpose,reg_date,reg_reason,debt_ratio,damaged_fee,situation_of_prepaid,subject_reg_order,secured_types_and_scope,secured_confirmation_date,ttl_amount_of_secured,certificate_no,right_person,rights_type,rights_value,rights_subject,rights_scope,assign_setting_restrictions,other_reg_items)
        values ('{pdf_nm}','{title}','{lot}','{print_dttm}','{land_item}',{type_seq},'{reg_order}','{delayed_interest}','{co_guarantee_area_no}','{co_guarantee_building_no}','{rent}','{paper_no}','{duration}','{receipted_year}','{address}','{interest_rate}','{instructions}','{other_guarantee_scope}','{property_by_payoff}','{settlement_date}','{id}','{set_purpose}','{reg_date}','{reg_reason}','{debt_ratio}','{damaged_fee}','{situation_of_prepaid}','{subject_reg_order}','{secured_types_and_scope}','{secured_confirmation_date}','{ttl_amount_of_secured}','{certificate_no}','{right_person}','{rights_type}','{rights_value}','{rights_subject}','{rights_scope}','{assign_setting_restrictions}','{other_reg_items}')
    """
    cursor2.execute(script)
    conn.commit()
    cursor2.close()
    conn.close()

# 取得所有檔案與子目錄名稱
files = listdir(path)
names = locals()
flist = []
for i in range(len(files)):
    if comnm in files[i] and doc in files[i]:
        flist.append(files[i])

# 以迴圈處理
pdfnm_list = []
for f in range(len(flist)):
    if '.csv' in flist[f]:
        pass
    else:        
        pdfnm = flist[f]
        pdfnm_list.append(pdfnm)

file_seq = 0
for f in range(len(pdfnm_list)):
    file_seq += 1
    print('Reader: ',file_seq)
    pdfnm = pdfnm_list[f]
    pdf_nm = pdfnm
    fp = open(rf'{mypath}\{pdfnm}','rb')
    parser = PDFParser(fp)
    document = PDFDocument(parser)
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    
    page_seq = 0
    pdfcnt = []
    
    for page in PDFPage.create_pages(document):
        page_seq += 1
        interpreter.process_page(page)
        layout = device.get_result()
        
        obj_seq = 0        
        pdfpkg = []
        for x in layout:
            obj_seq += 1
            if isinstance(x, LTTextBox):
                pdfword = (x.get_text().strip())
                pdfpkg = (f,file_seq,page_seq,obj_seq,pdfnm,pdfword)
                pdfcnt.append(pdfpkg)
    fp.close()
    
    #取標題、列印日期，頁面主要內容，雜訊清理
    elcontent = ''
    for i in range(len(pdfcnt)):
        
        if pdfcnt[i][2] == 1 and pdfcnt[i][3] == 1:
            title = pdfcnt[i][5]
        elif pdfcnt[i][2] == 1 and pdfcnt[i][3] == 2:
            lot = pdfcnt[i][5]
        elif pdfcnt[i][2] == 1 and pdfcnt[i][3] == 3:
            print_dttm = pdfcnt[i][5][5:30].strip()       
        else:
            data = (pdfcnt[i][0],pdfcnt[i][1],pdfcnt[i][2],pdfcnt[i][3],pdfcnt[i][4],pdfcnt[i][5])
            elcontent = elcontent + pdfcnt[i][5].replace('㈯','土').replace('㆞','地').replace('㈰','日').replace('㈪','月').replace('㆖','上').replace('㉂','自').replace('㈲','有').replace('㆟','人').replace('㆒','一').replace('㊠','項').replace('㈶','財').replace('㆗','中').replace('㈮','金').replace('㈴','名')

    section1 = '土地標示部'
    section2 = '土地所有權部'
    section3 = '土地他項權利部'
    
    # 依三種文件部別 歸類pdf物件
    if section1 in elcontent and section2 in elcontent and section3 in elcontent:
        land_label = ('land_label',elcontent[findchar(section1,elcontent,1):findchar(section2,elcontent,1)])
        land_own = ('land_own',elcontent[findchar(section2,elcontent,1):findchar(section3,elcontent,1)])
        land_other = ('land_other',elcontent[findchar(section3,elcontent,1):len(elcontent)])
    elif section1 in elcontent and section2 in elcontent:
        land_label = ('land_label',elcontent[findchar(section1,elcontent,1):findchar(section2,elcontent,1)])
        land_own = ('land_own',elcontent[findchar(section2,elcontent,1):len(elcontent)])
        land_other = ('land_other','')
    else:
        # 列印 spc ref 未歸納之內容
        AWCsv(exception,elcontent)

    land_data = []
    land_data.append(land_label)
    land_data.append(land_own)
    land_data.append(land_other)
    
    # 定義pdf內文 各種type 區分界線
    land_pargs = []
    for i in range(len(land_data)):
        item = land_data[i][0]
        item_spec = docspec(spec_ref,item,'')
        item_col_desc = list(item_spec['col_desc'])
        item_cols = len(item_col_desc)
        pars = land_data[i][1].count(item_col_desc[item_cols-1])
    
        for j in range(pars):
        
            char_e = item_col_desc[item_cols-1]
            
            if item == 'land_label' and j + 1 < pars:
                char_s = item_col_desc[0]
                len_s = findchar(char_s,land_data[i][1],1)
            
                char_e = item_col_desc[0]
                len_e = findchar(char_e,land_data[i][1],2)
        
            elif item == 'land_label' and j + 1 == pars:
                char_s = item_col_desc[0]
                len_s = findchar(char_s,land_data[i][1],1)
            
                char_e = ''
                len_e = len(land_data[i][1])
            
            
            elif item != 'land_label' and j + 1 < pars:
                char_s = rf'（{str((j+1)/10000+1/100000)[2:6]}）' + item_col_desc[0]
                len_s = findchar(char_s,land_data[i][1],1)
            
                char_e = rf'（{str((j+2)/10000+1/100000)[2:6]}）' + item_col_desc[0]
                len_e = findchar(char_e,land_data[i][1],1)
        
            else :
                char_s = rf'（{str((j+1)/10000+1/100000)[2:6]}）' + item_col_desc[0]
                char_e = ''
            
                len_s = findchar(char_s,land_data[i][1],1)
                len_e = len(land_data[i][1])
                
            land_parg = land_data[i][1][len_s-1:len_e]        
            parg_sub = (item,pars,j+1,char_s,len_s,char_e,len_e,land_parg)            
            land_pargs.append(parg_sub)
    
    # 建立pdf內文的切割規格      
    index = []        
    for k in range(len(land_pargs)):
        col_list = list(docspec(spec_ref,land_pargs[k][0],'')['col'])
        col_desc_list = list(docspec(spec_ref,land_pargs[k][0],'')['col_desc'])
        col_desc_list[0]=land_pargs[k][3]
        
        for l in range(len(col_list)):
            lens = land_pargs[k][7].find(col_desc_list[l])
            lens_d = lens/1000000+1/10000000
            lensc = "%.6f" % lens_d
            
            if lens > 0:
                index.append((land_pargs[k][0],land_pargs[k][1],land_pargs[k][2],col_list[l],col_desc_list[l],lensc,lens))
                
    index.sort(key=lambda index:index[0]+str(index[2])+index[5])
    
    # 依切割規格執行分段任務
    land_pargs_pd = pd.DataFrame(land_pargs).rename(columns={0:'item',1:'y1',2:'y2',3:'char_s',4:'len_s',5:'char_e',6:'len_e',7:'content'})
    index_pd = pd.DataFrame(index).rename(columns={0:'item',1:'x1',2:'x2',3:'col_name',4:'char_s'})
    item_label = []
    item_own = []
    item_other = []
    names = locals()
    for i in range(len(land_pargs_pd)):
        land_item = list(land_pargs_pd[i:i+1]['item'])[0]
        y1 = list(land_pargs_pd[i:i+1]['y1'])[0]
        type_seq = list(land_pargs_pd[i:i+1]['y2'])[0]
        ta_content = list(land_pargs_pd[i:i+1]['content'])[0]
    
        a = index_pd['item'] == land_item
        b = index_pd['x1'] == y1
        c = index_pd['x2'] == type_seq
        spec_index = index_pd[a & b & c]
    
        ### 欄位清空
        spec_all = docspec(spec_ref,'','')['col'].drop_duplicates().reset_index(drop=True)
        cols_all = len(spec_all)
        k = 0
        while k+1 <= cols_all:
            col_name = spec_all[k]
            names['%s' % col_name] = ''
            k += 1
    
        for j in range(len(spec_index)):
            col_name = list(spec_index[j:j+1]['col_name'])[0]
            char_s = list(spec_index[j:j+1]['char_s'])[0]
            ta_value_s = findchar(char_s,ta_content,1)+len(char_s)
            if j + 1 < len(spec_index):
                char_e = list(spec_index[j+1:j+2]['char_s'])[0]
                ta_value_e = findchar(char_e,ta_content,1)
            else:
                char_e = 'end'
                ta_value_e = len(ta_content)
            
            ta_value = ta_content[ta_value_s:ta_value_e].strip().replace('（續次頁）','')
            ta_value = str(' '.join(ta_value.split()))
        
            if '列印時間' in ta_value:
                ta_value = ta_value[0:findchar('列印時間',ta_value,1)]                
            
            if len(ta_value) > 4000:
                ta_value = ta_value[0:4000]
            
            names['%s' % col_name] = ta_value
        
            if land_item == 'land_label':
                item_label.append(names['%s' % col_name])
            elif land_item == 'land_own':
                item_own.append(names['%s' % col_name])
            elif land_item == 'land_other':
                item_other.append(names['%s' % col_name])
            else:
                pass
            
        if land_item == 'land_label':
            Insertlandlabel(pdf_nm,title,lot,print_dttm,land_item,type_seq,reg_date,reg_reason,area,use_partition,type_of_use,present_value,building_no,other_reg_items)
        elif land_item == 'land_own':
            Insertlandown(pdf_nm,title,lot,print_dttm,land_item,type_seq,reg_order,reg_date,reg_reason,cause_date,owner,id,rights_scope,title_no,declared_price,pretransfer_value_or_original_price,rights_obtained_scope,reg_other_of_other_right,other_reg_items)
        elif land_item == 'land_other':
            Insertlandother(pdf_nm,title,lot,print_dttm,land_item,type_seq,reg_order,delayed_interest,co_guarantee_area_no,co_guarantee_building_no,rent,paper_no,duration,receipted_year,address,interest_rate,instructions,other_guarantee_scope,property_by_payoff,settlement_date,id,set_purpose,reg_date,reg_reason,debt_ratio,damaged_fee,situation_of_prepaid,subject_reg_order,secured_types_and_scope,secured_confirmation_date,ttl_amount_of_secured,certificate_no,right_person,rights_type,rights_value,rights_subject,rights_scope,assign_setting_restrictions,other_reg_items)
        else:
            pass