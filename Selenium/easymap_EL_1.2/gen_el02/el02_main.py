# csv Package
import csv
import time
import argparse
import pandas as pd
import pymssql
import datetime as dt
import sys
import os

from etl_func import *
from batch_args import *
from config import *
c_index = pd.read_csv(r'C:\Py_Project\project\LandMoi\address_index.csv').drop_duplicates()

server = db['server']
database = db['database']
username = db['username']
password = db['password']
fromtb = db['fromtb']
totb = db['totb']
entitytype = db['entitytype']
batstoptime = dt.datetime.strptime(db['batstoptime'], '%Y-%m-%d %H:%M:%S.%f')

validtrans = 0
actiondttm = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
Statusdttm = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
Machine = hostname()
EntityPhase = 'el02_main.py'

obs = src_obs(server,username,password,database,fromtb,totb,entitytype)
print(obs)
datetime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())

new_collection(server,username,password,database,entitytype,fromtb,totb)
updatesql(server,username,password,database,obs,actiondttm,Statusdttm,'Done',entitytype,Machine,EntityPhase)
obs = src_obs(server,username,password,database,fromtb,totb,entitytype)
print(obs)
if obs == 0 :
    SUCC_mail(obs)
    pass
else:
    ERR_mail(obs)
    pass

        
