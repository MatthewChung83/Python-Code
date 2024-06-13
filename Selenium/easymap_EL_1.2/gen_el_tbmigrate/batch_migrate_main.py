import os
import datetime
from batch_args import *
from etl_func import *
from config import *

server = db['server']
database = db['database']
username = db['username']
password = db['password']
fromtb = db['fromtb']
Machine = hostname()

el07_info = dbfromel07(server,username,password,database,fromtb,Machine)

for i in range(len(el07_info)):
    print(el04migrate(server,username,password,database,'el04',el07_info[i][0]))
    print(el05migrate(server,username,password,database,'el05',el07_info[i][0]))