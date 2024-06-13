def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num
        
def src_obs(server,username,password,database,fromtb,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*) from [{fromtb}]
    where (status <> 'done' or status is null) and type = '{entitytype}'"""  
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]

def dbfrom(server,username,password,database,fromtb,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select * from [{fromtb}] where (status <> 'done' or status is null) and type = '{entitytype}'
    order by psid
    offset 0 row fetch next 1 rows only
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def toSQL(docs, totb, server, database, username, password):
    import pyodbc
    #conn_cmd = 'DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password
    conn_cmd = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password
    with pyodbc.connect(conn_cmd) as cnxn:
        cnxn.autocommit = False
        with cnxn.cursor() as cursor:
            data_keys = ','.join(docs[0].keys())
            data_symbols = ','.join(['?' for _ in range(len(docs[0].keys()))])
            insert_cmd = """INSERT INTO {} ({})
            VALUES ({})""".format(totb,data_keys,data_symbols )
            data_values = [tuple(doc.values()) for doc in docs]
            cursor.executemany( insert_cmd,data_values )
            cnxn.commit()
            
def EL03_etl(doc):
    EL03_result = []
    EL03_result.append({
        "eaid":doc[0],
        "pid":doc[1],
        "addressi":doc[2],
        "casei":doc[3],
        "type":doc[4],
        
        "Bd_District":doc[5],
        "Bd_GeoOffice":doc[6],
        "Bd_Lot":doc[7],
        "BuildNo":doc[8],
        "Bd_AreaSq":doc[9],
        "Bd_FloorNo":doc[10],
        "Bd_ByFloor":doc[11],
        "Bd_CompletionDate":doc[12],
        "Bd_Purpose":doc[13],
        
        "Ld_District":doc[14],
        "Ld_GeoOffice":doc[15],
        "Ld_Lot":doc[16],
        "LandNo":doc[17],
        "LandNo_f":doc[18],
        "LandNo_b":doc[19],
        "Ld_AreaSq":doc[20],
        "Ld_Value":doc[21],
        "Ld_Price":doc[22],
        
        "result":doc[23],
        "insertdate":doc[24],
        
        "bd_log":doc[25],
        "ld_log":doc[26], 
    })
    return EL03_result

def updatesql(server,username,password,database,status,info,INQCode,Taxreturnchannel,Taxreturnplat,taxreturndate,taxOffice,taxOfficeAddr,taxOfficePhone,entitytype,psid,pid):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    update taxreturntb 
    set info = '{info}' ,status = '{status}', INQCode = '{INQCode}', Taxreturnchannel = '{Taxreturnchannel}',Taxreturnplat='{Taxreturnplat}',
                taxreturndate='{taxreturndate}',taxOffice = '{taxOffice}',taxOfficeAddr = '{taxOfficeAddr}',taxOfficePhone = '{taxOfficePhone}'
    WHERE type='{entitytype}' and psid = '{psid}' and pid = '{pid}';
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    return script

def check_obs(server,username,password,database,fromtb,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*) from [{totb}]
    where type = '{entitytype}'"""    
    cursor.execute(script)
    obs = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return list(obs[0])[0]