# -*- coding: utf-8 -*-
"""
Created on Thu May 19 16:52:22 2022

@author: admin
"""


# -*- coding: utf-8 -*-
"""
Created on Thu May 19 15:11:31 2022

@author: matthew5043
"""
import pyodbc
import pandas as pd

from datetime import datetime, date, timedelta
today = date.today() + timedelta(days = 0)
today1 = str(today).replace('-','')


def casei(server,username,password,database,i,today):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select	distinct r.casei ,bk.bankname from treasure.skiptrace.dbo.returntb r
 
left join treasure.skiptrace.dbo.clienttb ct on r.Client = ct.clienti
left join treasure.skiptrace.dbo.banktb bk on  bk.Masterclienti = ct.MasterClientI

where r.requestdate >= '{today}' and
 r.Client ='{i}' 
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def ID(server,username,password,database,i,today):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select	distinct r.ID_Number ,bk.bankname from treasure.skiptrace.dbo.returntb r
 
    left join treasure.skiptrace.dbo.clienttb ct on r.Client = ct.clienti
    left join treasure.skiptrace.dbo.banktb bk on  bk.Masterclienti = ct.MasterClientI

    where r.requestdate > ='{today}' and
    r.Client ='{i}' 
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src
def note(server,username,password,database,today):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select	distinct 
	r.casei,
	case --
	when n.Note like'%'+ p.Phone+'%' then replace(n.note,p.Phone,'*')
		 when n.note like'%'+ r.ID_Number +'%' then replace(n.note,substring(r.ID_Number,2,9),'*')
		 when patindex( '%[鎮|市]區%' ,n.note) > 0 and charindex( '號' ,n.note) > 0 then replace(n.note,replace(replace(substring(n.note,1,patindex('%號%',n.note)-1),(SELECT substring(substring(n.note,1,patindex('%號%',n.note)-1), 1,patindex('%[縣|市]%',substring(n.note,1,patindex('%號%',n.note)-1)))),''),(SELECT substring(replace(substring(n.note,1,patindex('%號%',n.note)-1),(SELECT substring(substring(n.note,1,patindex('%號%',n.note)-1), 1,patindex('%[縣|市]%',substring(n.note,1,patindex('%號%',n.note)-1)))),''), 1,patindex('%[鎮|市]區%',(select replace(substring(n.note,1,patindex('%號%',n.note)-1),(SELECT substring(substring(n.note,1,patindex('%號%',n.note)-1), 1,patindex('%[縣|市]%',substring(n.note,1,patindex('%號%',n.note)-1)))),''))))),''),'*****')
		 when patindex( '%[鄉|鎮|區|縣|市]%' ,n.note) > 0 and charindex( '號' ,n.note) > 0 then replace(n.note,replace(replace(substring(n.note,1,patindex('%號%',n.note)-1),(SELECT substring(substring(n.note,1,patindex('%號%',n.note)-1), 1,patindex('%[縣|市]%',substring(n.note,1,patindex('%號%',n.note)-1)))),''),(SELECT substring(replace(substring(n.note,1,patindex('%號%',n.note)-1),(SELECT substring(substring(n.note,1,patindex('%號%',n.note)-1), 1,patindex('%[縣|市]%',substring(n.note,1,patindex('%號%',n.note)-1)))),''), 1,patindex('%[鄉|鎮|區|縣|市]%',(select replace(substring(n.note,1,patindex('%號%',n.note)-1),(SELECT substring(substring(n.note,1,patindex('%號%',n.note)-1), 1,patindex('%[縣|市]%',substring(n.note,1,patindex('%號%',n.note)-1)))),''))))),''),'*****')
		 when patindex( '%第%' ,n.note) > 0 and patindex( '%號%' ,n.note) > 0 then replace(n.note,replace(replace(n.note,substring (n.note,1,patindex( '%第%' ,n.note) ),''),'號',''),'****')
		 when patindex( '%段%' ,n.note) > 0 and patindex( '%地號%' ,n.note) > 0 then replace(n.note,replace(replace(n.note,substring (n.note,1,patindex( '%段%' ,n.note) ),''),'地號',''),'****')
		 when n.note like '%電話%' then ''
		 when n.note like '%' + ps.name +'%' then replace(n.note,ps.name,'***')
		 when n.note like '%' + ps.OName +'%' then replace(n.note,ps.OName,'***')
		 else n.Note
		 End as Note  ,
	bk.bankname 
	into #test1
	from treasure.skiptrace.dbo.returntb r
 
	left join treasure.skiptrace.dbo.notetb n on r.casei = n.casei 
    left join treasure.skiptrace.dbo.casepersontb cp on r.CaseI = cp.CaseI 
	left join treasure.skiptrace.dbo.persontb ps on cp.PersonI = ps.PersonI
	left join treasure.skiptrace.dbo.PersonPhoneTb pp on cp.PersonI = pp.PersonI
	left join treasure.skiptrace.dbo.phonetb p on p.PhoneI = pp.PhoneI
	left join treasure.skiptrace.dbo.PersonAddrTb pa on cp.PersonI = pa.PersonI
	left join treasure.skiptrace.dbo.AddressV adr on adr.AddressI = pa.AddressI
    left join treasure.skiptrace.dbo.clienttb ct on r.Client = ct.clienti
    left join treasure.skiptrace.dbo.banktb bk on  bk.Masterclienti = ct.MasterClientI
    
    where r.requestdate >='{today}' and n.NoteType <> '0'  order by CaseI
    select distinct
        n.CaseI,
        case  when n.note like '%' + ps.name +'%' then replace(n.note,ps.name,'***')
	          when n.note like '%' + ps.OName +'%' then replace(n.note,ps.OName,'***')
              when n.Note like'%'+ p.Phone+'%' then ''
	          when replace(replace(replace(replace(n.Note,'(',''),')',''),' ',''),'__','') like'%'+ replace(replace(replace(replace(p.Phone,'(',''),')',''),' ',''),'__','')  +'%' then replace(n.note,p.Phone,replace(replace(replace(replace(p.Phone,'(',''),')',''),' ',''),substring(replace(replace(replace(p.Phone,'(',''),')',''),' ',''),4,12),'*'))
	          End as note,
      bankname
      into #test2
      from #test1 n
      left join treasure.skiptrace.dbo.casepersontb cp on n.CaseI = cp.CaseI 
      left join treasure.skiptrace.dbo.persontb ps on cp.PersonI = ps.PersonI
      left join treasure.skiptrace.dbo.PersonPhoneTb pp on cp.PersonI = pp.PersonI
      left join treasure.skiptrace.dbo.phonetb p on p.PhoneI = pp.PhoneI
    create table #test3
    (casei int,note nvarchar(max),bankname nvarchar(500))
    insert into #test3
    select distinct
        n.CaseI,
        case  
	          when n.Note like'%'+p.Phonetype+'%' then replace(n.note,p.Phonetype,'')
	          End as note,
      bankname
      from #test2 n
      left join treasure.skiptrace.dbo.casepersontb cp on n.CaseI = cp.CaseI 
      left join treasure.skiptrace.dbo.persontb ps on cp.PersonI = ps.PersonI
      left join treasure.skiptrace.dbo.PersonPhoneTb pp on cp.PersonI = pp.PersonI
      left join treasure.skiptrace.dbo.phonetb p on p.PhoneI = pp.PhoneI
      where n.note is not null
     select * from #test3
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src


def note_obs(server,username,password,database,today):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select	distinct 
	count(*)
	
	from treasure.skiptrace.dbo.returntb r
 
	left join treasure.skiptrace.dbo.notetb n on r.casei = n.casei 
    left join treasure.skiptrace.dbo.casepersontb cp on r.CaseI = cp.CaseI 
	left join treasure.skiptrace.dbo.PersonPhoneTb pp on cp.PersonI = pp.PersonI
	left join treasure.skiptrace.dbo.phonetb p on p.PhoneI = pp.PhoneI
	left join treasure.skiptrace.dbo.PersonAddrTb pa on cp.PersonI = pa.PersonI
	left join treasure.skiptrace.dbo.AddressV adr on adr.AddressI = pa.AddressI
    left join treasure.skiptrace.dbo.clienttb ct on r.Client = ct.clienti
    left join treasure.skiptrace.dbo.banktb bk on  bk.Masterclienti = ct.MasterClientI
    
    where r.requestdate >='{today}' 
    
    
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]

def phone(server,username,password,database,today):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select distinct r.casei,p.Phone ,bk.bankname from treasure.skiptrace.dbo.returntb r
    left join treasure.skiptrace.dbo.casepersontb cp on r.CaseI = cp.CaseI 
	left join treasure.skiptrace.dbo.PersonPhoneTb pp on cp.PersonI = pp.PersonI
	left join treasure.skiptrace.dbo.phonetb p on p.PhoneI = pp.PhoneI
    left join treasure.skiptrace.dbo.clienttb ct on r.Client = ct.clienti
    left join treasure.skiptrace.dbo.banktb bk on  bk.Masterclienti = ct.MasterClientI
    
    where r.requestdate >='{today}' and p.phone is not null 
     
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def address(server,username,password,database,today):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select distinct r.casei,adr.City,adr.Town ,bk.bankname from treasure.skiptrace.dbo.returntb r
    left join treasure.skiptrace.dbo.casepersontb cp on r.CaseI = cp.CaseI 
	
	left join treasure.skiptrace.dbo.PersonAddrTb pp on cp.PersonI = pp.PersonI
	left join treasure.skiptrace.dbo.AddressV adr on adr.AddressI = pp.AddressI
    left join treasure.skiptrace.dbo.clienttb ct on r.Client = ct.clienti
    left join treasure.skiptrace.dbo.banktb bk on  bk.Masterclienti = ct.MasterClientI
    
    where r.requestdate >='{today}'  and  adr.City is not null 
     
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

#取 client
server = 'vnt07.ucs.com' 
database = 'UIS'
username = 'pyuser' 
password = 'Ucredit7607'  
cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
query = "select	distinct str(Client) as client from treasure.skiptrace.dbo.returntb where requestdate > ="+"'"+str(today)+"'"
df = pd.read_sql(query, cnxn)
client = df['client'].tolist()

#取 folder 銀行別
shortname = []
for i in client:
    
    server = 'vnt07.ucs.com' 
    database = 'UIS'
    username = 'pyuser' 
    password = 'Ucredit7607'  
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
    cursor = cnxn.cursor()
    query = """select bankname from treasure.skiptrace.dbo.banktb bk 
    left join treasure.skiptrace.dbo.clienttb ct on  bk.Masterclienti = ct.MasterClientI 
    where ct.ClientI =  """+"'"+str(i)+"'"""
    df = pd.read_sql(query, cnxn)
    folder = df['bankname'].tolist()
    shortname.append(folder)

#pirnt(shortname)

#建立資料夾
import os 
for i in shortname :
    try:
        os.mkdir(rf'C:\Users\admin\Desktop\test\test\{i[0]}')
    except:
        pass
    try:
        os.mkdir(rf'C:\Users\admin\Desktop\test\test\{i[0]}\{today1}')
    except:
        pass

#Casei 刪除作業
import re
for i in client:
   i =str(i)
   src = casei(server,username,password,database,str(i),str(today))[0]
   log = 'INFO : '+str(today)+' delete '+' '+str(src[0])+'\n'
   path =rf'C:\Users\admin\Desktop\test\test\{src[1]}\{today1}\\'+str(today)+'_'+'DELETE_CASE_FILE.txt'
   f = open(path,'a')
   f.writelines(log)
   f.close()
   print(src[0])
   print(src[1])

#ID 刪除作業
import re
for i in client:
   i =str(i)
   src = ID(server,username,password,database,str(i),str(today))[0]
   log = 'INFO : '+str(today)+' delete '+' '+ src[0][0]+re.sub('[0-9]', '*',src[0][1:6])+src[0][7:10]+'\n'
   path =rf'C:\Users\admin\Desktop\test\test\{src[1]}\{today1}\\'+str(today)+'_'+'DELETE_ID_FILE.txt'
   f = open(path,'a')
   f.writelines(log)
   f.close()
   print(src[0])
   print(src[1])


#Note 刪除作業
src = note(server,username,password,database,str(today))
#print(src)
for i in src:
   try:
       print(i)
       #if '(' in str(i[1]) and ')' in str(i[1]) :
       #    str(i[1]).replace('1','*').replace('2','*').replace('3','*').replace('4','*').replace('5','*').replace('6','*').replace('7','*').replace('8','*').replace('9','*')
       #elif '-' in i[1] :
       #    str(i[1]).replace('1','*').replace('2','*').replace('3','*').replace('4','*').replace('5','*').replace('6','*').replace('7','*').replace('8','*').replace('9','*')
       if '09' in i[1] or '01'in i[1] or '02'in i[1] or '03'in i[1] or '04'in i[1] or '05'in i[1] or '06'in i[1] or '07'in i[1] or '08'in i[1] :
           str(i[1]).replace('1','*').replace('2','*').replace('3','*').replace('4','*').replace('5','*').replace('6','*').replace('7','*').replace('8','*').replace('9','*')
        
     
       log = 'INFO : '+str(today)+' delete '+str(i[0])+' '+ str(i[1])+'\n'
       path =rf'C:\Users\admin\Desktop\test\test\{i[2]}\{today1}\\'+str(today)+'_'+'DELETE_NOTE_FILE.txt'
       f = open(path,'a')
       f.writelines(log)
       f.close()
       #print(src[0])
       #print(src[1])
   except:
       pass

#電話刪除作業
src_phone = phone(server,username,password,database,str(today))
import re
for i in src_phone:

   phone_dtl = i[1].replace('(','').replace(')','').replace(' ','')
   log = 'INFO : '+str(today)+' delete '+' '+ phone_dtl[0:4]+re.sub('[0-9]', '*',phone_dtl[4:10])+'\n'
   path =rf'C:\Users\admin\Desktop\test\test\{i[2]}\{today1}\\'+str(today)+'_'+'DELETE_Phone_FILE.txt'
   f = open(path,'a')
   f.writelines(log)
   f.close()
   #print(src[0])
   #print(src[1])

#地址除作業
#print(address[0:address.find('區','')+1]+'*******')
src_address = address(server,username,password,database,str(today))
import re
for i in src_address:
   #i =str(i)
   #src = address(server,username,password,database,str(today))[0]
   
   log = 'INFO : '+str(today)+' delete '+' '+ str(i[1])+str(i[2])+'*************號'+'\n'
   path =rf'C:\Users\admin\Desktop\test\test\{i[3]}\{today1}\\'+str(today)+'_'+'DELETE_ADDR_FILE.txt'
   f = open(path,'a')
   f.writelines(log)
   f.close()
    
    
    
    
    
    
    
    
    
    
    
    
    