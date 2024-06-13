# -*- coding: utf-8 -*-
"""
Created on Wed Jan 11 17:01:14 2023

@author: matthew5043
"""


import datetime
import time
import requests
import sys
import json
import pymssql


import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


from bs4 import BeautifulSoup
#from fake_useragent import UserAgent
from requests import Session




def delete(server,username,password,database,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    delete from  [{totb}]
    
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()

def toSQL(docs, totb, server, database, username, password):
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

def mail():
    import smtplib
    from email.mime.text import MIMEText
    from email.header import Header
    sender = 'matthew5043@ucs.com'
    receivers = ['matthew5043@ucs.com'] # 接收郵件
    # 三個引數：第一個為文字內容，第二個 plain 設定文字格式，第三個 utf-8 設定編碼
    message = MIMEText('Clockin_records end', 'plain', 'utf-8')
    message['From'] = Header('matthew5043', 'utf-8') # 傳送者
    message['To'] = Header('matthew5043', 'utf-8') # 接收者
    subject = 'Clockin_records end'
    message['Subject'] = Header(subject, 'utf-8')
    try:
        smtpObj = smtplib.SMTP('vrh19.ucs.com')
        smtpObj.sendmail(sender, receivers, message.as_string())
        print('郵件傳送成功')
    except smtplib.SMTPException:
        print('Error: 無法傳送郵件')
        
def Clockin_records(doc):
    Clockin_records= []

    Clockin_records.append({
        
        'SYS_ROWID':doc[0], 
        'SYS_COMPANYID':doc[1],
        'TMP_DECCOMPANYID':doc[2],
        'TMP_DECCOMPANYNAME':doc[3],
        'TMP_DECCOMPANYENGNAME':doc[4],
        'DEPARTID':doc[5],
        'DEPARTID2':doc[6],
        'DEPARTNAME':doc[7],
        'DEPARTENGNAME':doc[8],
        'SERIAL':doc[9],
        'PROFITID':doc[10],
        'PROFITNAME':doc[11],
        'TMP_EMPLOYEEID':doc[12],
        'TMP_EMPLOYEENAME':doc[13],
        'TMP_WORKID':doc[14],
        'TMP_WORKNAME':doc[15],
        'STARTTIME':doc[16],
        'ENDTIME':doc[17],
        'ATTENDDATE':doc[18],
        'WEEKDAY':doc[19],
        'CARDNO':doc[20],
        'WORKTYPE':doc[21],
        'PREARRIVETIME':doc[22],
        'PRELATEMINS':doc[23],
        'BOVERTIME':doc[24],
        'BOVERTIMESTATUS':doc[25],
        'BOFFOVERTIME':doc[26],
        'BOFFOVERTIMESTATUS':doc[27],
        'WORKTIME':doc[28],
        'WORKTIMESTATUS':doc[29],
        'STATUS':doc[30],
        'OFFWORKTIME':doc[31],
        'OFFWORKTIMESTATUS':doc[32],
        'STATUS2':doc[33],
        'AOVERTIME':doc[34],
        'AOVERTIMESTATUS':doc[35],
        'AOFFOVERTIME':doc[36],
        'AOFFOVERTIMESTATUS':doc[37],
        'SWORKHOURS':doc[38],
        'REALWORKMINUTES':doc[39],
        'REALWORKHOURS':doc[40],
        'LEAVEHOURS':doc[41],
        'OFFLEAVEHOURS':doc[42],
        'OVERHOURS':doc[43],
        'TOTALHOURS':doc[44],
        'DIFFHOURS':doc[45],
        'NOTE':doc[46],
        'ATTENDDATES':doc[47],
        'ISTATUS':doc[48],
        'ISTATUS2':doc[49],
        'EMPLOYEEID':doc[50],
        'WORKID':doc[51],
        'DWORKTIME':doc[52],
        'DOFFWORKTIME':doc[53],
        'GD1':doc[54],
        'GD2':doc[55],
        'GD3':doc[56],
        'GD4':doc[57],
        'GD5':doc[58],
        'GD6':doc[59],
        'LEAVESTARTTIME':doc[60],
        'LEAVEENDTIME':doc[61],
        'LEAVENAME':doc[62],
        'OVERSTARTTIME':doc[63],
        'OVERENDTIME':doc[64],
        'LEAVEID':doc[65],
        'OVERID':doc[66],
        'OVERTYPE':doc[67],
        'DOOVERTYPE':doc[68],
        'LATEMINS':doc[69],
        'EARLYMINS':doc[70],
        'FORGETTIMES':doc[71],
        'VACATIONTYPEID':doc[72],
        'GPSLOCATION':doc[73],
        'SWNOTE':doc[74],
        'GPSADDRESS':doc[75],
        'GPSLOCATION2':doc[76],
        'SWNOTE2':doc[77],
        'GPSADDRESS2':doc[78],
        'IPADDRESS':doc[79],
        'IPADDRESS2':doc[80],
        'SOURCETYPE':doc[81],
        'SOURCETYPE2':doc[82],
        'JOBCODEID':doc[83],
        'JOBCODENAME':doc[84],
        'JOBCODE2ID':doc[85],
        'JOBCODE2NAME':doc[86],
        'JOBLEVELID':doc[87],
        'JOBLEVELNAME':doc[88],
        'JOBRANKID':doc[89],
        'JOBRANKNAME':doc[90],
        'JOBTYPEID':doc[91],
        'JOBTYPENAME':doc[92],
        'JOBCATEGORYID':doc[93],
        'JOBCATEGORYNAME':doc[94],
        'HASATTENDSUM':doc[95],
        'JOBSTATUS2':doc[96],
        'PRELATETIMES':doc[97],
        'LATETIMES':doc[98],
        'EARLYTIMES':doc[99],
        'insertdate':doc[100],
    })
    return Clockin_records       
db = {
    'server': 'RICHES',
    'database': 'CL_Daily',
    'username': 'CLUSER',
    'password': 'Ucredit7607',
    'totb':'Clockin_records',

    
}
wbinfo = {
    #'main_url':'https://scsservices.azurewebsites.net/api/systemobject/',
    #'api_url':'https://scsservices.azurewebsites.net/api/businessobject/',
    #'main_url':'https://client.scshr.com/api/businessobject/',
    #'api_url':'https://client.scshr.com/api/businessobject/',
    'main_url':'https://hr.ucs.com/SCSRwd/api/systemobject/',
    'api_url':'https://hr.ucs.com/SCSRwd/api/businessobject/',
}
server,database,username,password,totb= db['server'],db['database'],db['username'],db['password'],db['totb']
main_url,api_url = wbinfo['main_url'],wbinfo['api_url']



getdate = datetime.date.today()

getdate = str(getdate).replace('-','/')

print(getdate)

print(f'打卡記錄同步-起始時間: {getdate}')

SenssionID=''        
headers = {'Content-type': 'application/json'}
data={
		"Action": "Login",
		"SessionGuid": "",
		"Value":{
			"$type": "AIS.Define.TLogingInputArgs, AIS.Define",
			"CompanyID": "scs164",
			"UserID": "api",
			"Password": "api$1234",
			"LanguageID": "zh-TW"			
		}
}
data_json = json.dumps(data)
response = requests.post(main_url, data=data_json, headers=headers,verify=False)
result = response.json()

if result.get('Result'):
	SenssionID = result.get('SessionGuid')
else :
	print(result.get('Result'),result.get('Message'))
delete(server,username,password,database,totb)

if SenssionID !="":
    # 提取請假單-get tmp_agentempi
    #getdate = datetime.datetime.now().strftime("%Y/%m/%d")
    #getdate = (datetime.datetime.now()+datetime.timedelta(days=-1)).strftime("%Y/%m/%d")
    #tomorrow = (datetime.datetime.now()+datetime.timedelta(days=1)).strftime("%Y/%m/%d")
    #getdate = '2022/10/21'
    data = {
            "Action": "ExecReport",
            "SessionGuid": SenssionID,
            "ProgID": "RATT017",
            "Value": {
            "$type": "AIS.Define.TExecReportInputArgs, AIS.Define",
            "UIType": "Report",
            "ReportID": "",
            "ReportTailID": "",
            "FilterItems": [
              {
                "$type": "AIS.Define.TFilterItem, AIS.Define",
                "FieldName": "@COMPANYID",
                "FilterValue": "SCS164"
              },
              {
                "$type": "AIS.Define.TFilterItem, AIS.Define",
                "FieldName": "@STARTDATE",
                "FilterValue": getdate
              },
              {
                "$type": "AIS.Define.TFilterItem, AIS.Define",
                "FieldName": "@ENDDATE",
                "FilterValue": getdate
              }
            ],
            "UserFilter": ""
          }
         }
    data_json = json.dumps(data)
    response = requests.post(api_url, data=data_json, headers=headers,verify=False)
    result =response.json()
    #print(result)
    datatype='DataSet'
    data = result.get(datatype).get('ReportBody')

    for i in range((len(data))):
        
        SYS_ROWID = data[i].get('SYS_ROWID')
        SYS_COMPANYID = data[i].get('SYS_COMPANYID')
        TMP_DECCOMPANYID = data[i].get('TMP_DECCOMPANYID')
        TMP_DECCOMPANYNAME = data[i].get('TMP_DECCOMPANYNAME')
        TMP_DECCOMPANYENGNAME = data[i].get('TMP_DECCOMPANYENGNAME')
        DEPARTID = data[i].get('DEPARTID')
        DEPARTID2 = data[i].get('DEPARTID2')
        DEPARTNAME = data[i].get('DEPARTNAME')
        DEPARTENGNAME = data[i].get('DEPARTENGNAME')
        SERIAL = data[i].get('SERIAL') 
        PROFITID = data[i].get('PROFITID')
        PROFITNAME = data[i].get('PROFITNAME')
        TMP_EMPLOYEEID = data[i].get('TMP_EMPLOYEEID')
        TMP_EMPLOYEENAME = data[i].get('TMP_EMPLOYEENAME')
        TMP_WORKID = data[i].get('TMP_WORKID')
        TMP_WORKNAME = data[i].get('TMP_WORKNAME')
        STARTTIME = data[i].get('STARTTIME')
        ENDTIME = data[i].get('ENDTIME') 
        ATTENDDATE = data[i].get('ATTENDDATE')
        WEEKDAY = data[i].get('WEEKDAY') 
        CARDNO = data[i].get('CARDNO')
        WORKTYPE = data[i].get('WORKTYPE')
        PREARRIVETIME = data[i].get('PREARRIVETIME') 
        PRELATEMINS = data[i].get('PRELATEMINS') 
        BOVERTIME = data[i].get('BOVERTIME') 
        BOVERTIMESTATUS = data[i].get('BOVERTIMESTATUS') 
        BOFFOVERTIME = data[i].get('BOFFOVERTIME')
        BOFFOVERTIMESTATUS = data[i].get('BOFFOVERTIMESTATUS')
        WORKTIME = data[i].get('WORKTIME') 
        WORKTIMESTATUS = data[i].get('WORKTIMESTATUS')
        STATUS = data[i].get('STATUS') 
        OFFWORKTIME = data[i].get('OFFWORKTIME')
        OFFWORKTIMESTATUS = data[i].get('OFFWORKTIMESTATUS') 
        STATUS2 = data[i].get('STATUS2')
        AOVERTIME = data[i].get('AOVERTIME')
        AOVERTIMESTATUS = data[i].get('AOVERTIMESTATUS')
        AOFFOVERTIME = data[i].get('AOFFOVERTIME')
        AOFFOVERTIMESTATUS = data[i].get('AOFFOVERTIMESTATUS')
        SWORKHOURS = data[i].get('SWORKHOURS')
        REALWORKMINUTES = data[i].get('REALWORKMINUTES')
        REALWORKHOURS = data[i].get('REALWORKHOURS')
        LEAVEHOURS = data[i].get('LEAVEHOURS') 
        OFFLEAVEHOURS = data[i].get('OFFLEAVEHOURS')
        OVERHOURS = data[i].get('OVERHOURS')
        TOTALHOURS = data[i].get('TOTALHOURS')
        DIFFHOURS = data[i].get('DIFFHOURS') 
        NOTE = data[i].get('NOTE') 
        ATTENDDATES = data[i].get('ATTENDDATES')
        ISTATUS = data[i].get('ISTATUS') 
        ISTATUS2 = data[i].get('ISTATUS2')
        EMPLOYEEID = data[i].get('EMPLOYEEID') 
        WORKID = data[i].get('WORKID')
        DWORKTIME = data[i].get('DWORKTIME')
        DOFFWORKTIME = data[i].get('DOFFWORKTIME') 
        GD1 = data[i].get('GD1') 
        GD2 = data[i].get('GD2') 
        GD3 = data[i].get('GD3')
        GD4 = data[i].get('GD4') 
        GD5 = data[i].get('GD5')
        GD6 = data[i].get('GD6') 
        LEAVESTARTTIME = data[i].get('LEAVESTARTTIME')
        LEAVEENDTIME = data[i].get('LEAVEENDTIME')
        LEAVENAME = data[i].get('LEAVENAME')
        OVERSTARTTIME = data[i].get('OVERSTARTTIME')
        OVERENDTIME = data[i].get('OVERENDTIME')
        LEAVEID = data[i].get('LEAVEID') 
        OVERID = data[i].get('OVERID')
        OVERTYPE = data[i].get('OVERTYPE') 
        DOOVERTYPE = data[i].get('DOOVERTYPE')
        LATEMINS = data[i].get('LATEMINS') 
        EARLYMINS = data[i].get('EARLYMINS') 
        FORGETTIMES = data[i].get('FORGETTIMES')
        VACATIONTYPEID = data[i].get('VACATIONTYPEID')
        GPSLOCATION = data[i].get('GPSLOCATION')
        SWNOTE = data[i].get('SWNOTE')
        GPSADDRESS = data[i].get('GPSADDRESS')
        GPSLOCATION2 = data[i].get('GPSLOCATION2') 
        SWNOTE2 = data[i].get('SWNOTE2')
        GPSADDRESS2 = data[i].get('GPSADDRESS2')
        IPADDRESS = data[i].get('IPADDRESS')
        IPADDRESS2 = data[i].get('IPADDRESS2') 
        SOURCETYPE = data[i].get('SOURCETYPE')
        SOURCETYPE2 = data[i].get('SOURCETYPE2')
        JOBCODEID = data[i].get('JOBCODEID')
        JOBCODENAME = data[i].get('JOBCODENAME') 
        JOBCODE2ID = data[i].get('JOBCODE2ID')
        JOBCODE2NAME = data[i].get('JOBCODE2NAME')
        JOBLEVELID = data[i].get('JOBLEVELID')
        JOBLEVELNAME = data[i].get('JOBLEVELNAME') 
        JOBRANKID = data[i].get('JOBRANKID')
        JOBRANKNAME = data[i].get('JOBRANKNAME')
        JOBTYPEID = data[i].get('JOBTYPEID')
        JOBTYPENAME = data[i].get('JOBTYPENAME') 
        JOBCATEGORYID = data[i].get('JOBCATEGORYID')
        JOBCATEGORYNAME = data[i].get('JOBCATEGORYNAME')
        HASATTENDSUM = data[i].get('HASATTENDSUM')
        JOBSTATUS2 = data[i].get('JOBSTATUS2') 
        PRELATETIMES = data[i].get('PRELATETIMES')
        LATETIMES = data[i].get('LATETIMES')
        EARLYTIMES = data[i].get('EARLYTIMES')

        insertdate = getdate
        update_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        docs = (SYS_ROWID,SYS_COMPANYID,TMP_DECCOMPANYID,TMP_DECCOMPANYNAME,TMP_DECCOMPANYENGNAME,
            DEPARTID,DEPARTID2,DEPARTNAME,DEPARTENGNAME,SERIAL,PROFITID,PROFITNAME,
            TMP_EMPLOYEEID,TMP_EMPLOYEENAME,TMP_WORKID,TMP_WORKNAME,STARTTIME,ENDTIME,
            ATTENDDATE,WEEKDAY,CARDNO,WORKTYPE,PREARRIVETIME,PRELATEMINS,BOVERTIME,BOVERTIMESTATUS,
            BOFFOVERTIME,BOFFOVERTIMESTATUS,WORKTIME,WORKTIMESTATUS,STATUS,OFFWORKTIME,OFFWORKTIMESTATUS,
            STATUS2,AOVERTIME,AOVERTIMESTATUS,AOFFOVERTIME,AOFFOVERTIMESTATUS,SWORKHOURS,REALWORKMINUTES,
            REALWORKHOURS,LEAVEHOURS,OFFLEAVEHOURS,OVERHOURS,TOTALHOURS,DIFFHOURS,NOTE,ATTENDDATES,
            ISTATUS,ISTATUS2,EMPLOYEEID,WORKID,DWORKTIME,DOFFWORKTIME,GD1,GD2,GD3,GD4,GD5,GD6,
            LEAVESTARTTIME,LEAVEENDTIME,LEAVENAME,OVERSTARTTIME,OVERENDTIME,LEAVEID,OVERID,
            OVERTYPE,DOOVERTYPE,LATEMINS,EARLYMINS,FORGETTIMES,VACATIONTYPEID,GPSLOCATION,SWNOTE,
            GPSADDRESS,GPSLOCATION2,SWNOTE2,GPSADDRESS2,IPADDRESS,IPADDRESS2,SOURCETYPE,SOURCETYPE2,
            JOBCODEID,JOBCODENAME,JOBCODE2ID,JOBCODE2NAME,JOBLEVELID,JOBLEVELNAME,JOBRANKID,JOBRANKNAME,
            JOBTYPEID,JOBTYPENAME,JOBCATEGORYID,JOBCATEGORYNAME,HASATTENDSUM,JOBSTATUS2,PRELATETIMES,LATETIMES,
            EARLYTIMES,update_date)
        
        Clockin_records_result = Clockin_records(docs)
        #print(empleavetb_result)
        toSQL(Clockin_records_result, totb, server, database, username, password)

mail()















