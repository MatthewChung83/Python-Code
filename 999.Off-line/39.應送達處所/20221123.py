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
    select count(*) from [20221123_name]

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

    
    select *from [20221123_name] 
    where score_full =''
    order by 編號,ADDRGIS
    
	offset 0 row fetch next 1 rows only
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def update(server,username,password,database,addressi,num,Score_full,Score_County,Score_Town,Score_Village,Score_Nebrh ,Score_Road ,Score_section ,Score_Area ,Score_Lane ,Score_Alley ,Score_SubAlley ,Score_Num ,Score_Floor ,Score_Room):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update [20221123_name]
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
    where addressi = '{addressi}' and 編號 = '{num}'
    
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
    num = src[0]
    GIS_TRANS = src[6]
    GIS_TRANS_County= src[7]
    GIS_TRANS_Town= src[8]
    GIS_TRANS_Village= src[9]
    GIS_TRANS_Nebrh= src[10]
    GIS_TRANS_Road= src[11]
    GIS_TRANS_Section= src[12]
    GIS_TRANS_Area= src[13]
    GIS_TRANS_Lane= src[14]
    GIS_TRANS_Alley= src[15]
    GIS_TRANS_SubAlley= src[16]
    GIS_TRANS_Num= src[17]
    GIS_TRANS_Floor= src[18]
    GIS_TRANS_Room= src[19]
    AddrGIS = src[20]
    AddrGIS_County= src[21]
    AddrGIS_Town= src[22]
    AddrGIS_Village= src[23]
    AddrGIS_Nebrh= src[24]
    AddrGIS_Road= src[25]
    AddrGIS_Section= src[26]
    AddrGIS_Area= src[27]
    AddrGIS_Lane= src[28]
    AddrGIS_Alley= src[29]
    AddrGIS_SubAlley= src[30]
    AddrGIS_Num= src[31]
    AddrGIS_Floor= src[32]
    AddrGIS_Room= src[33]
    addressi = src[34]
    #full address score
    Score_full = fuzz.ratio(GIS_TRANS,AddrGIS)
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
    update(server,username,password,database,addressi,num,Score_full,Score_County,Score_Town,Score_Village,Score_Nebrh ,Score_Road ,Score_section ,Score_Area ,Score_Lane ,Score_Alley ,Score_SubAlley ,Score_Num ,Score_Floor ,Score_Room)




#%%
compareScore = fuzz.ratio('花蓮縣壽豐鄉鹽寮村007鄰字爵德157之3號','花蓮縣壽豐鄉鹽寮村００７鄰福德１５７之３號')
print(compareScore)






































