# -*- coding: utf-8 -*-
"""
Created on Mon May  9 18:28:59 2022

@author: admin
"""


def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num
def src_obs(server,username,password,database,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*) from {totb} where [出生年月日] is null and [姓名] is not null
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]
def votetw(doc):
    votetw= []

    votetw.append({
        
        "選舉別":doc[0],
        "選舉區":doc[1],
        "登記日期":doc[2],
        "姓名":doc[3],        
        "推薦之政黨":doc[4],
        "出生年月日":doc[5],
        "選舉區1":doc[6],
        "現任":doc[7],
    })
    return votetw



def dbfrom(server,username,password,database,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select 
    rowid, 
    選舉別,
    選舉區_市 = (SELECT substring([選舉區], 1,patindex('%[縣|市]%',[選舉區]))),
    選舉區_區 = case when  [選舉區] like '%選舉區%' and 選舉別 <> '市民代表' then ''
    				when [選舉區] like '%鎮區%'then (SELECT substring(replace([選舉區],(SELECT substring([選舉區], 1,patindex('%[縣|市]%',[選舉區]))),''), 1,patindex('%[鎮|市]區%',(select replace([選舉區],(SELECT substring([選舉區], 1,patindex('%[縣|市]%',[選舉區]))),'')))))+'區'
    				when [選舉區] like '%市區%'then (SELECT substring(replace([選舉區],(SELECT substring([選舉區], 1,patindex('%[縣|市]%',[選舉區]))),''), 1,patindex('%[鎮|市]區%',(select replace([選舉區],(SELECT substring([選舉區], 1,patindex('%[縣|市]%',[選舉區]))),'')))))+'區'
    			else (SELECT substring(replace([選舉區],(SELECT substring([選舉區], 1,patindex('%[縣|市]%',[選舉區]))),''), 1,patindex('%[鄉|鎮|區|縣|市]%',(select replace([選舉區],(SELECT substring([選舉區], 1,patindex('%[縣|市]%',[選舉區]))),''))))) end,
    選舉區_選區 = replace(replace(replace(replace(選舉區,(SELECT substring([選舉區], 1,patindex('%[縣|市]%',[選舉區]))),''),case when  [選舉區] like '%選舉區%' and 選舉別 <> '市民代表' then '' when [選舉區] like '%鎮區%'then (SELECT substring(replace([選舉區],(SELECT substring([選舉區], 1,patindex('%[縣|市]%',[選舉區]))),''), 1,patindex('%[鎮|市]區%',(select replace([選舉區],(SELECT substring([選舉區], 1,patindex('%[縣|市]%',[選舉區]))),'')))))+'區'
    						when [選舉區] like '%市區%'then (SELECT substring(replace([選舉區],(SELECT substring([選舉區], 1,patindex('%[縣|市]%',[選舉區]))),''), 1,patindex('%[鎮|市]區%',(select replace([選舉區],(SELECT substring([選舉區], 1,patindex('%[縣|市]%',[選舉區]))),'')))))+'區'
    			else (SELECT substring(replace([選舉區],(SELECT substring([選舉區], 1,patindex('%[縣|市]%',[選舉區]))),''), 1,patindex('%[鄉|鎮|區|縣|市]%',(select replace([選舉區],(SELECT substring([選舉區], 1,patindex('%[縣|市]%',[選舉區]))),''))))) end,''),CHAR(10),''),' ',''),
    登記日期,
    姓名,
    推薦之政黨
     from UIS.dbo.votetw_02  where [出生年月日] is null and [姓名] is not null order by rowid
	offset 0 row fetch next 1 rows only
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def update(server,username,password,database,totb,birthday,area_req,vote_n,rowid,name_href,area_check):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update [{totb}]
    set [出生年月日] = '{birthday}',[選舉區1] = N'{area_req}',[現任] = N'{vote_n}',[姓名1] = N'{name_href}',[區域符合] = N'{area_check}'
    where rowid = '{rowid}'
    
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
