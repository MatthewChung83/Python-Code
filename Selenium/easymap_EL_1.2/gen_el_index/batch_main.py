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
batstrtime = datetime.datetime.now()
batstoptime = batstrtime + datetime.timedelta(minutes=240)

el02_index = dbfromel02(server,username,password,database,fromtb,Machine)

for i in range(len(el02_index)):
    entitytype = el02_index[i][0]
    path = el02_index[i][3]
    file = el02_index[i][2]
    script = 'python '+path+file+' "{}" "{}"'.format(entitytype,batstoptime)
    print(script)
    os.system('{}'.format(script))

el03_index = dbfromel03(server,username,password,database,fromtb,Machine)

for j in range(len(el03_index)):
    entitytype = el03_index[j][0]
    path = el03_index[j][3]
    file = el03_index[j][2]
    script = 'python '+path+file+' "{}" "{}"'.format(entitytype,batstoptime)
    print(script)
    os.system('{}'.format(script))
print(Machine)