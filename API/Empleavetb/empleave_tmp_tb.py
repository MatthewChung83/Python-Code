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



getdate = datetime.date.today()
#last_month = getdate.replace(month=getdate.month - 1)
getdate = str(getdate).replace('-','/')
#last_month = str(last_month).replace('-','/')
print(getdate)
#print(last_month)
print(f'請假記錄同步-起始時間: {getdate}')

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
      "ProgID": "RATT004",
      "Value": {
        "$type": "AIS.Define.TExecReportInputArgs, AIS.Define",
        "UIType": "Report",
        "ReportID": "",
        "ReportTailID": "",
        "FilterItems": [
          {
            "$type": "AIS.Define.TFilterItem, AIS.Define",
            "FieldName": "A.SYS_CompanyID",
            "FilterValue": "SCS164"
          },
          #{
          #  "$type": "AIS.Define.TFilterItem, AIS.Define",
          #  "FieldName": "C.SYS_Date",
          #  "FilterValue": getdate,
          #  "ComparisonOperator": "GreaterOrEqual"
          #},
          {
            "$type": "AIS.Define.TFilterItem, AIS.Define",
            "FieldName": "IsNull(DT.StartDate, D.StartDate)",
            "FilterValue": getdate,
            #"ComparisonOperator": "GreaterOrEqual"
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
        
        LEAVETYPE = data[i].get('LEAVETYPE')
        SYS_COMPANYID = data[i].get('SYS_COMPANYID')
        TMP_DECCOMPANYID = data[i].get('TMP_DECCOMPANYID')
        TMP_DECCOMPANYNAME = data[i].get('TMP_DECCOMPANYNAME')
        TMP_DECCOMPANYENGNAME = data[i].get('TMP_DECCOMPANYENGNAME')
        DEPARTID = data[i].get('DEPARTID')
        DEPARTID2 = data[i].get('DEPARTID2')
        DEPARTNAME = data[i].get('DEPARTNAME')
        DEPARTENGNAME = data[i].get('DEPARTENGNAME')
        EMPLOYEEID = data[i].get('EMPLOYEEID') 
        EMPLOYEENAME = data[i].get('EMPLOYEENAME')
        SYS_ENGNAME = data[i].get('SYS_ENGNAME')
        SEX = data[i].get('SEX')
        SYS_VIEWID = data[i].get('SYS_VIEWID')
        SYS_DATE = data[i].get('SYS_DATE')
        VACATIONID = data[i].get('VACATIONID')
        VACATIONNAME = data[i].get('VACATIONNAME')
        VACATIONENGNAME = data[i].get('VACATIONENGNAME') 
        SVACATIONID = data[i].get('SVACATIONID')
        SVACATIONNAME = data[i].get('SVACATIONNAME') 
        SVACATIONENGNAME = data[i].get('SVACATIONENGNAME')
        STARTDATE = data[i].get('STARTDATE')
        STARTTIME = data[i].get('STARTTIME') 
        ENDDATE = data[i].get('ENDDATE') 
        ENDTIME = data[i].get('ENDTIME') 
        LEAVEDAYS = data[i].get('LEAVEDAYS') 
        LEAVEHOURS = data[i].get('LEAVEHOURS')
        LEAVEMINUTES = data[i].get('LEAVEMINUTES')
        HOURWAGES = data[i].get('HOURWAGES') 
        LEAVEMONEY = data[i].get('LEAVEMONEY')
        AGENTID = data[i].get('AGENTID') 
        AGENTNAME = data[i].get('AGENTNAME')
        MAINNOTE = data[i].get('MAINNOTE') 
        SUBNOTE = data[i].get('SUBNOTE')
        SYS_FLOWFORMSTATUS = data[i].get('SYS_FLOWFORMSTATUS')
        OFFLEAVEDAYS = data[i].get('OFFLEAVEDAYS')
        OFFLEAVEHOURS = data[i].get('OFFLEAVEHOURS')
        OFFLEAVEMINUTES = data[i].get('OFFLEAVEMINUTES')
        REALLEAVEDAYS = data[i].get('REALLEAVEDAYS')
        REALLEAVEHOURS = data[i].get('REALLEAVEHOURS')
        REALLEAVEMINUTES = data[i].get('REALLEAVEMINUTES')
        CUTDATE = data[i].get('CUTDATE') 
        SPECIALDATE = data[i].get('SPECIALDATE')
        STARGETNAME = data[i].get('STARGETNAME')
        SENDDATE = data[i].get('SENDDATE')
        SOURCETAG = data[i].get('SOURCETAG') 
        OUTSIDENAME = data[i].get('OUTSIDENAME') 
        OUTSIDETEL = data[i].get('OUTSIDETEL')
        ISLEAVE = data[i].get('ISLEAVE') 
        ISCOMEBACK = data[i].get('ISCOMEBACK')
        EMPTEL = data[i].get('EMPTEL') 
        RESTPLACE = data[i].get('RESTPLACE')
        EMPADDRESS = data[i].get('EMPADDRESS')
        NOTE2 = data[i].get('NOTE2')
        PRJOECTID = data[i].get('PRJOECTID')
        TMP_PRJOECTID = data[i].get('TMP_PRJOECTID')
        TMP_PRJOECTNAME = data[i].get('TMP_PRJOECTNAME')
        TMP_PRJOECTENGNAME = data[i].get('TMP_PRJOECTENGNAME')
        DIRECTID = data[i].get('DIRECTID')
        TMP_DIRECTID = data[i].get('TMP_DIRECTID') 
        PMANAGERID = data[i].get('PMANAGERID')
        TMP_PMANAGERID = data[i].get('TMP_PMANAGERID') 
        APPROVER3ID = data[i].get('APPROVER3ID')
        TMP_APPROVER3ID = data[i].get('TMP_APPROVER3ID') 
        APPROVER4ID = data[i].get('APPROVER4ID') 
        TMP_APPROVER4ID = data[i].get('TMP_APPROVER4ID') 
        GD1 = data[i].get('GD1') 
        GD2 = data[i].get('GD2') 
        GD3 = data[i].get('GD3')
        GD4 = data[i].get('GD4') 
        GD5 = data[i].get('GD5')
        GD6 = data[i].get('GD6') 
        VACATIONTYPEID = data[i].get('VACATIONTYPEID')
        VACATIONTYPENAME = data[i].get('VACATIONTYPENAME')
        VACATIONTYPEENGNAME = data[i].get('VACATIONTYPEENGNAME')
        CDEPARTID = data[i].get('CDEPARTID')
        CDEPARTNAME = data[i].get('CDEPARTNAME') 
        CDEPARTENGNAME = data[i].get('CDEPARTENGNAME')
        insertdate = getdate
        update_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        docs = (LEAVETYPE,SYS_COMPANYID,TMP_DECCOMPANYID,TMP_DECCOMPANYNAME,
                TMP_DECCOMPANYENGNAME,DEPARTID,DEPARTID2,DEPARTNAME,DEPARTENGNAME,
                EMPLOYEEID, EMPLOYEENAME, SYS_ENGNAME, SEX, SYS_VIEWID, SYS_DATE, 
                VACATIONID, VACATIONNAME, VACATIONENGNAME, SVACATIONID, SVACATIONNAME, 
                SVACATIONENGNAME, STARTDATE, STARTTIME, ENDDATE, ENDTIME, LEAVEDAYS, 
                LEAVEHOURS, LEAVEMINUTES, HOURWAGES, LEAVEMONEY, AGENTID, AGENTNAME, 
                MAINNOTE, SUBNOTE, SYS_FLOWFORMSTATUS, OFFLEAVEDAYS, OFFLEAVEHOURS, 
                OFFLEAVEMINUTES, REALLEAVEDAYS, REALLEAVEHOURS, REALLEAVEMINUTES, 
                CUTDATE, SPECIALDATE, STARGETNAME, SENDDATE, SOURCETAG, OUTSIDENAME, 
                OUTSIDETEL, ISLEAVE, ISCOMEBACK, EMPTEL, RESTPLACE, EMPADDRESS, 
                NOTE2, PRJOECTID, TMP_PRJOECTID, TMP_PRJOECTNAME, TMP_PRJOECTENGNAME, 
                DIRECTID, TMP_DIRECTID, PMANAGERID, TMP_PMANAGERID, APPROVER3ID, 
                TMP_APPROVER3ID, APPROVER4ID, TMP_APPROVER4ID, GD1, GD2, GD3, 
                GD4, GD5, GD6, VACATIONTYPEID, VACATIONTYPENAME, VACATIONTYPEENGNAME, 
                CDEPARTID, CDEPARTNAME, CDEPARTENGNAME,insertdate,update_date)
        
        empleave_tmp_tb_result = empleave_tmp_tb(docs)
        #print(empleavetb_result)
        toSQL(empleave_tmp_tb_result, totb, server, database, username, password)

mail()















