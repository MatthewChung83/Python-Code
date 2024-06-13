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

el05_index = dbfromel05(server,username,password,database,fromtb,Machine)

for i in range(len(el05_index)):
    entitytype = el05_index[i][0]
    path = r' C:\Py_Project\project\easymap_EL_1.2\gen_el05\\'
    file = 'el05_main.py'
    script = 'python '+path+file+' "{}" '.format(entitytype)
    #print(script)
    os.system('{}'.format(script))