def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num
def src_obs(server,username,password,database,fromtb,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    
    select (select count(*) from {fromtb} b
	left join [{totb}] r on b.ID = r.ID 
	where r.ID is null )
	+
	(select count(*) from {fromtb} b
	left join [{totb}] r on b.ID = r.ID 
	where (select DATEDIFF(mm, update_date, getdate()))> = 3 )   
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]
def tmnewa(doc):
    tmnewa= []

    tmnewa.append({
        "Name":doc[0],
        "ID":doc[1],
        "birthday":doc[2],
        "insurance_type":doc[3],
        "insurance_num":doc[4],
        "insurance_status":doc[5],
        "insurance_query_type":doc[6],
        "insurance_query_num":doc[7],
        "insurance_query_status":doc[8],
        "update_date":doc[9],
        "note":doc[10],
    })
    return tmnewa
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


def dbfrom(server,username,password,database,fromtb,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select  b.ID,b.name,b.m,r.rowid,r.insurance_type,r.insurance_num,r.insurance_query_type,r.insurance_query_num,b.flg
    into #test
    from {fromtb} b
	left join [{totb}] r on b.ID = r.ID 
	where r.ID is null 
	insert into #test
	select  b.ID,b.name,b.m,r.rowid,r.insurance_type,r.insurance_num,r.insurance_query_type,r.insurance_query_num,b.flg
    from {fromtb} b
	left join [{totb}] r on b.ID = r.ID 
	where (select DATEDIFF(mm, update_date, getdate()))> = 3   
    select * from #test order by flg desc
    --order by personi
	offset 0 row fetch next 1 rows only
    
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def update(server,username,password,database,totb,note,ID,updatetime,rowid,Insurance_type_web,Insurance_num_web,Insurance_status_web,Insurance_query_type_web,ID_query_web,Insurance_query_num_web,Insurance_query_status_web):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update [{totb}]
    set note = '{note}',Insurance_type = '{Insurance_type_web}',Insurance_num = '{Insurance_num_web}',Insurance_status = '{Insurance_status_web}',
        Insurance_query_type = '{Insurance_query_type_web}',Insurance_query_num = '{Insurance_query_num_web}',Insurance_query_status ='{Insurance_query_status_web}' , update_date = '{updatetime}'
    where id = '{ID}' and rowid = '{rowid}'
    
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
def exit_obs(server,username,password,database,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count (distinct ID)
    from [{totb}]
    where  update_date > = convert(varchar(10),getdate(),111)
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]