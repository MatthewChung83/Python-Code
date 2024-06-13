# -*- coding: utf-8 -*-
"""
Created on Mon May  9 18:28:59 2022

@author: admin
"""
#import argparse

#parser = argparse.ArgumentParser(description='外部參數')
#parser.add_argument('type', type=str , help='名單名稱')      # 名單名稱

#args = parser.parse_args()
#entitytype = args.type


db = {
    'server': 'RICHES',
    'database': 'CL_Daily',
    'username': 'CLUSER',
    'password': 'Ucredit7607',
    'totb':'empleave_tmp_tb',

    
}
wbinfo = {
    #'main_url':'https://scsservices.azurewebsites.net/api/systemobject/',
    #'api_url':'https://scsservices.azurewebsites.net/api/businessobject/',
    #'main_url':'https://client.scshr.com/api/businessobject/',
    #'api_url':'https://client.scshr.com/api/businessobject/',
    'main_url':'https://hr.ucs.com/SCSRwd/api/systemobject/',
    'api_url':'https://hr.ucs.com/SCSRwd/api/businessobject/',
}