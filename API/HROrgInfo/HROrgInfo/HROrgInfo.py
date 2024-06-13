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
              "Action": "Find",
              "SessionGuid": SenssionID,
              "ProgID": "HUM0010300",
              "Value": {
                "$type": "AIS.Define.TFindInputArgs, AIS.Define",
                "SelectCount": 500,
                "SelectFields": "SYS_VIEWID,SYS_NAME,SYS_ENGNAME,TMP_PDEPARTID,TMP_PDEPARTNAME,TMP_MANAGERID,TMP_MANAGERNAME",
                "SystemFilterOptions": "Session,DataPermission,EmployeeLevel",
                "IsBuildSelectedField": 'true',
                "IsBuildFlowLightSignalField": 'true'
              }
            }
    data_json = json.dumps(data)
    response = requests.post(api_url, data=data_json, headers=headers,verify=False)
    result =response.json()
    datatype='DataTable'
    data = result.get(datatype)

    for i in range((len(data))):
        SYS_VIEWID = data[i].get('SYS_VIEWID')
        SYS_NAME = data[i].get('SYS_NAME')
        SYS_ENGNAME = data[i].get('SYS_ENGNAME')
        TMP_PDEPARTID = data[i].get('TMP_PDEPARTID')
        TMP_PDEPARTNAME = data[i].get('TMP_PDEPARTNAME')
        TMP_PDEPARTENGNAME = data[i].get('TMP_PDEPARTENGNAME')
        SYS_ID = data[i].get('SYS_ID')
        TMP_MANAGERID = data[i].get('TMP_MANAGERID')
        TMP_MANAGERNAME = data[i].get('TMP_MANAGERNAME')
        TMP_MANAGERENGNAME = data[i].get('TMP_MANAGERENGNAME')




        docs = (SYS_VIEWID,SYS_NAME,SYS_ENGNAME,TMP_PDEPARTID,TMP_PDEPARTNAME,
                TMP_PDEPARTENGNAME,SYS_ID,TMP_MANAGERID,TMP_MANAGERNAME,TMP_MANAGERENGNAME,)
        NewCash_result = NewCash(docs)
        toSQL(NewCash_result, totb, server, database, username, password)
        continue

mail()














