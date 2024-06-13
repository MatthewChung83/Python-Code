# -*- coding: utf-8 -*-
"""
Created on Tue Nov 22 15:18:50 2022

@author: admin
"""


def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num
def src_obs(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*) from [FL_GIS_reply]where score_full is null

    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]

def dbfrom(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""

    
    select *from [FL_GIS_reply] 
    where score_full is null
    order by ID
    
	offset 0 row fetch next 1 rows only
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def update(server,username,password,database,ID,Score_full,Score_County,Score_Town,Score_Village,Score_Nebrh ,Score_Road ,Score_section ,Score_Area ,Score_Lane ,Score_Alley ,Score_SubAlley ,Score_Num ,Score_Floor ,Score_Room):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update [FL_GIS_reply]
    set Score_full = '{Score_full}',
    Score_County='{Score_County}',
    Score_Town='{Score_Town}' ,
    Score_Village='{Score_Village}' ,
    Score_Nebrh='{Score_Nebrh}',
    Score_Road='{Score_Road}' ,
    Score_section='{Score_section}' ,
    Score_Area='{Score_Area}' ,
    Score_Lane='{Score_Lane}' ,
    Score_Alley='{Score_Alley}' ,
    Score_SubAlley='{Score_SubAlley}' ,
    Score_Num='{Score_Num}' ,
    Score_Floor='{Score_Floor}' ,
    Score_Room='{Score_Room}' 
    where ID = '{ID}'
    
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
db = {
    'server': 'vnt07.ucs.com',
    'database': 'uis',
    'username': 'pyuser',
    'password': 'Ucredit7607',   
}

   
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

server,database,username,password = db['server'],db['database'],db['username'],db['password']
obs = src_obs(server,username,password,database)

for i in foo(-1,obs-1):   
    src = dbfrom(server,username,password,database)[0]
    ID = src[0]
    GIS_TRANS = src[3]
    GIS_TRANS_County= src[4]
    GIS_TRANS_Town= src[5]
    GIS_TRANS_Village= src[6]
    GIS_TRANS_Nebrh= src[7]
    GIS_TRANS_Road= src[8]
    GIS_TRANS_Section= src[9]
    GIS_TRANS_Area= src[10]
    GIS_TRANS_Lane= src[11]
    GIS_TRANS_Alley= src[12]
    GIS_TRANS_SubAlley= src[13]
    GIS_TRANS_Num= src[14]
    GIS_TRANS_Floor= src[15]
    GIS_TRANS_Room= src[16]
    AddrGIS = src[18]
    AddrGIS_County= src[19]
    AddrGIS_Town= src[20]
    AddrGIS_Village= src[21]
    AddrGIS_Nebrh= src[22]
    AddrGIS_Road= src[23]
    AddrGIS_Section= src[24]
    AddrGIS_Area= src[25]
    AddrGIS_Lane= src[26]
    AddrGIS_Alley= src[27]
    AddrGIS_SubAlley= src[28]
    AddrGIS_Num= src[29]
    AddrGIS_Floor= src[30]
    AddrGIS_Room= src[31]
    
    #full address score
    Score_full = fuzz.ratio(GIS_TRANS,AddrGIS)
    print(GIS_TRANS,AddrGIS)
    #Score_County
    Score_County = fuzz.ratio(GIS_TRANS_County,AddrGIS_County)
    #Score_Town
    Score_Town = fuzz.ratio(GIS_TRANS_Town,AddrGIS_Town)
    #Score_Village
    Score_Village = fuzz.ratio(GIS_TRANS_Village,AddrGIS_Village)
    #Score_Nebrh
    Score_Nebrh = fuzz.ratio(GIS_TRANS_Nebrh,AddrGIS_Nebrh)
    #Score_Road
    Score_Road  = fuzz.ratio(GIS_TRANS_Road,AddrGIS_Road)
    #Score_section
    Score_section = fuzz.ratio(GIS_TRANS_Section,AddrGIS_Section)
    #Score_Area
    Score_Area = fuzz.ratio(GIS_TRANS_Area,AddrGIS_Area)
    #Score_Lane
    Score_Lane = fuzz.ratio(GIS_TRANS_Lane,AddrGIS_Lane)
    #Score_Alley
    Score_Alley = fuzz.ratio(GIS_TRANS_Alley,AddrGIS_Alley)
    #Score_SubAlley
    Score_SubAlley = fuzz.ratio(GIS_TRANS_SubAlley,AddrGIS_SubAlley)
    #Score_Num
    Score_Num = fuzz.ratio(GIS_TRANS_Num,AddrGIS_Num)
    #Score_Floor
    Score_Floor = fuzz.ratio(GIS_TRANS_Floor,AddrGIS_Floor)
    #Score_Room
    Score_Room = fuzz.ratio(GIS_TRANS_Room,AddrGIS_Room)
    update(server,username,password,database,ID,Score_full,Score_County,Score_Town,Score_Village,Score_Nebrh ,Score_Road ,Score_section ,Score_Area ,Score_Lane ,Score_Alley ,Score_SubAlley ,Score_Num ,Score_Floor ,Score_Room)




#%%
compareScore = fuzz.ratio('花蓮縣壽豐鄉鹽寮村007鄰字爵德157之3號','花蓮縣壽豐鄉鹽寮村7鄰福德157-3號')
print(compareScore)






































