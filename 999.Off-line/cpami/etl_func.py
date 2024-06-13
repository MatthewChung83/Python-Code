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
	(select count(*) from [{totb}]
	where (select DATEDIFF(mm, update_date, getdate()))> = 3 )   
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]
def cpami(doc):
    cpami= []

    cpami.append({
        "ID":doc[0],
        "Name":doc[1],
        "notation":doc[2],
        "architect_Name":doc[3],
        "license_num":doc[4],
        "certificate_num":doc[5],
        "office_name":doc[6],
        "update_date":doc[7],
        "note":doc[8],
    })
    return cpami
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
    select  b.ID,b.name,r.rowid,r.notation,r.architect_Name,r.license_num,r.certificate_num,r.office_name,b.flg
    into #test
    from {fromtb} b
	left join [{totb}] r on b.ID = r.ID 
	where r.ID is null 
	insert into #test
	select  r.ID,r.name,r.rowid,r.notation,r.architect_Name,r.license_num,r.certificate_num,r.office_name,'0' as flg
    from [{totb}] r
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

def update(server,username,password,database,totb,ID,notation_web,architect_Name_web,license_num_web,certificate_num_web,office_name_web,updatetime,note,rowid):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update [{totb}]
    set note = '{note}',notation = '{notation_web}',architect_Name = '{architect_Name_web}',license_num = '{license_num_web}',
        certificate_num = '{certificate_num_web}',office_name = '{office_name_web}', update_date = '{updatetime}'
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