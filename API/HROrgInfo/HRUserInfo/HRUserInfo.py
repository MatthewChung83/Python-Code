# -*- coding: utf-8 -*-
"""
Created on Thu Jan 20 22:42:02 2022

@author: admin
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



from config import *
from etl_func import *
from dict import *

server,database,username,password,totb= db['server'],db['database'],db['username'],db['password'],db['totb']
main_url,api_url = wbinfo['main_url'],wbinfo['api_url']



getdate = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
delete(server,username,password,database,totb)
print(f'人事資料表-起始時間: {getdate}')

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

if SenssionID !="":
    # 提取請假單-get tmp_agentempi
    #getdate = datetime.datetime.now().strftime("%Y/%m/%d")
    #getdate = (datetime.datetime.now()+datetime.timedelta(days=-1)).strftime("%Y/%m/%d")
    #tomorrow = (datetime.datetime.now()+datetime.timedelta(days=1)).strftime("%Y/%m/%d")
    #getdate = '2022/10/21'
    data = {  
      "Action": "ExecReport",
      "SessionGuid": SenssionID,
      "ProgID": "RHUM002",
      "Value": {
        "$type": "AIS.Define.TExecReportInputArgs, AIS.Define",
        "UIType": "Report",
        "ReportID": "",
        "ReportTailID": "",
        "FilterItems": "",
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
        SYS_ID = data[i].get('SYS_ID')
        SYS_COMPANYID = data[i].get('SYS_COMPANYID')
        TMP_DECCOMPANYID = data[i].get('TMP_DECCOMPANYID')
        TMP_DECCOMPANYNAME = data[i].get('TMP_DECCOMPANYNAME')
        TMP_DECCOMPANYENGNAME = data[i].get('TMP_DECCOMPANYENGNAME')
        DEPARTID = data[i].get('DEPARTID')
        DEPARTID2 = data[i].get('DEPARTID2')
        TMP_DEPARTNAME = data[i].get('TMP_DEPARTNAME')
        TMP_DEPARTENGNAME = data[i].get('TMP_DEPARTENGNAME')
        SERIAL = data[i].get('SERIAL')
        EMPLOYEEID2 = data[i].get('EMPLOYEEID2')
        EMPLOYEEID = data[i].get('EMPLOYEEID')
        EMPLOYEENAME = data[i].get('EMPLOYEENAME')
        EMPLOYEEENGNAME = data[i].get('EMPLOYEEENGNAME')
        IDNO = data[i].get('IDNO')
        JOBSTATUS = data[i].get('JOBSTATUS')
        COUNTRYNAME = data[i].get('COUNTRYNAME')
        BIRTHDATE = data[i].get('BIRTHDATE')
        SEX = data[i].get('SEX')
        MARRAGE = data[i].get('MARRAGE')
        BLOODTYPE = data[i].get('BLOODTYPE')
        BIRTHPLACE = data[i].get('BIRTHPLACE')
        OTHER_BIRTHPLACE = data[i].get('OTHER_BIRTHPLACE')
        VTITLEDEPARTID = data[i].get('VTITLEDEPARTID')
        HDEGREE = data[i].get('HDEGREE')
        ISTOP = data[i].get('ISTOP')
        STARTDATE = data[i].get('STARTDATE')
        WORKINGYEARS = data[i].get('WORKINGYEARS')
        SSTARTDATE = data[i].get('SSTARTDATE')
        SPECIALYEARS = data[i].get('SPECIALYEARS')
        GARRIVEDATE = data[i].get('GARRIVEDATE')
        MOIBLE = data[i].get('MOIBLE')
        OFFICETEL1 = data[i].get('OFFICETEL1')
        OFFICETEL2 = data[i].get('OFFICETEL2')
        PSNEMAI = data[i].get('PSNEMAI')
        EMAIL1 = data[i].get('EMAIL1')
        EMAIL2 = data[i].get('EMAIL2')
        REGTEL = data[i].get('REGTEL')
        REGADDRESS = data[i].get('REGADDRESS')
        COMMTEL = data[i].get('COMMTEL')
        COMMADDRESS = data[i].get('COMMADDRESS')
        EMERGENCYNAME = data[i].get('EMERGENCYNAME')
        EMERGENCYTELNO = data[i].get('EMERGENCYTELNO')
        EMERGENCYMOBILE = data[i].get('EMERGENCYMOBILE')
        EMERGENCYSEX = data[i].get('EMERGENCYSEX')
        TMP_EMERGENCYID = data[i].get('TMP_EMERGENCYID')
        TMP_EMERGENCYNAME = data[i].get('TMP_EMERGENCYNAME')
        JOBCODENAME = data[i].get('JOBCODENAME')
        JOBCODEENGNAME = data[i].get('JOBCODEENGNAME')
        JOBLEVELNAME = data[i].get('JOBLEVELNAME')
        JOBLEVELENGNAME = data[i].get('JOBLEVELENGNAME')
        JOBRANKNAME = data[i].get('JOBRANKNAME')
        JOBRANKENGNAME = data[i].get('JOBRANKENGNAME')
        JOBCODEID = data[i].get('JOBCODEID')
        JOBLEVELID = data[i].get('JOBLEVELID')
        JOBRANKID = data[i].get('JOBRANKID')
        GD1 = data[i].get('GD1')
        GD2 = data[i].get('GD2')
        GD3 = data[i].get('GD3')
        GD4 = data[i].get('GD4')
        GD5 = data[i].get('GD5')
        GD6 = data[i].get('GD6')
        GD1_ID = data[i].get('GD1_ID')
        GD2_ID = data[i].get('GD2_ID')
        SELFDEF1 = data[i].get('SELFDEF1')
        SELFDEF2 = data[i].get('SELFDEF2')
        SELFDEF3 = data[i].get('SELFDEF3')
        SELFDEF4 = data[i].get('SELFDEF4')
        SELFDEF5 = data[i].get('SELFDEF5')
        TMP_PROFITID = data[i].get('TMP_PROFITID')
        TMP_PROFITNAME = data[i].get('TMP_PROFITNAME')
        TMP_IDYCLASSID = data[i].get('TMP_IDYCLASSID')
        TMP_IDYCLASSNAME = data[i].get('TMP_IDYCLASSNAME')
        ISDIRECT = data[i].get('ISDIRECT')
        ETHNICID = data[i].get('ETHNICID')
        ISDISABILITY = data[i].get('ISDISABILITY')
        DISABILITYDEGREE = data[i].get('DISABILITYDEGREE')
        NOTE = data[i].get('NOTE')
        WORKINGYEARSYMD = data[i].get('WORKINGYEARSYMD')
        SELFDEF6 = data[i].get('SELFDEF6')
        SELFDEF7 = data[i].get('SELFDEF7')
        SELFDEF8 = data[i].get('SELFDEF8')
        JOBDESCRIPTION = data[i].get('JOBDESCRIPTION')
        JOBCONTRACT = data[i].get('JOBCONTRACT')



        docs = (SYS_ID,SYS_COMPANYID,TMP_DECCOMPANYID,TMP_DECCOMPANYNAME,TMP_DECCOMPANYENGNAME,DEPARTID,DEPARTID2,
                TMP_DEPARTNAME,TMP_DEPARTENGNAME,SERIAL,EMPLOYEEID2,EMPLOYEEID,EMPLOYEENAME,EMPLOYEEENGNAME,IDNO,
                JOBSTATUS,COUNTRYNAME,BIRTHDATE,SEX,MARRAGE,BLOODTYPE,BIRTHPLACE,OTHER_BIRTHPLACE,VTITLEDEPARTID,
                HDEGREE,ISTOP,STARTDATE,WORKINGYEARS,SSTARTDATE,SPECIALYEARS,GARRIVEDATE,MOIBLE,OFFICETEL1,OFFICETEL2,
                PSNEMAI,EMAIL1,EMAIL2,REGTEL,REGADDRESS,COMMTEL,COMMADDRESS,EMERGENCYNAME,EMERGENCYTELNO,EMERGENCYMOBILE,
                EMERGENCYSEX,TMP_EMERGENCYID,TMP_EMERGENCYNAME,JOBCODENAME,JOBCODEENGNAME,JOBLEVELNAME,JOBLEVELENGNAME,
                JOBRANKNAME,JOBRANKENGNAME,JOBCODEID,JOBLEVELID,JOBRANKID,GD1,GD2,GD3,GD4,GD5,GD6,GD1_ID,GD2_ID,SELFDEF1,
                SELFDEF2,SELFDEF3,SELFDEF4,SELFDEF5,TMP_PROFITID,TMP_PROFITNAME,TMP_IDYCLASSID,TMP_IDYCLASSNAME,ISDIRECT,
                ETHNICID,ISDISABILITY,DISABILITYDEGREE,NOTE,WORKINGYEARSYMD,SELFDEF6,SELFDEF7,SELFDEF8,JOBDESCRIPTION,
                JOBCONTRACT)
        NewCash_result = NewCash(docs)
        toSQL(NewCash_result, totb, server, database, username, password)
        continue

mail()














