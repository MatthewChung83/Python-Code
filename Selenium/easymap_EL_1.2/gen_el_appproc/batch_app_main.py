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

el03dtl_index = dbfromel03dtl(server,username,password,database,fromtb,Machine)
#batstrtime = datetime.datetime.now()
#batstoptime = batstrtime + datetime.timedelta(minutes=240)

for i in range(len(el03dtl_index)):
    entitytype = el03dtl_index[i][0]
    path = r' C:\Py_Project\project\easymap_EL_1.2\gen_el04\\'
    file = 'el04_main.py'
    script = 'python '+path+file+' "{}" '.format(entitytype)
    os.system('{}'.format(script))