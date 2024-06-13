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
	where r.ID is null)
	--+
	--(select count(*) from [{totb}] 
    --where (select DATEDIFF(mm, update_date, getdate()))> = 3 )    
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]
def travelagency(doc):
    travelagency= []

    travelagency.append({
        "ID":doc[0],
        "Name":doc[1],
        "Name_web":doc[2],
        "sex":doc[3],
        "company":doc[4],
        "company_type":doc[5],
        "emp_type":doc[6],
        "tranning_type":doc[7],
        "update_date":doc[8],
        "note":doc[9],
    })
    return travelagency
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
    select  b.ID,b.name,r.rowid,r.name_web,r.sex,r.company,r.company_type,r.emp_type,r.tranning_type,b.flg
    into #test
    from {fromtb} b
	left join [{totb}] r on b.ID = r.ID 
	where r.ID is null  
	--insert into #test
	--select  ID,name,rowid,Name_web,sex,company,company_type,emp_type,tranning_type,'' as flg
    --from [{totb}]
	--where (select DATEDIFF(mm, update_date, getdate()))> = 3   
    select * from #test order by flg desc,id
    --order by personi
	offset 0 row fetch next 1 rows only
    
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src

def update(server,username,password,database,totb,ID,Name_web,sex_web,company_web,company_type_web,emp_type_web,tranning_type_web,updatetime,note,rowid):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
        
    script = f"""
    update [{totb}]
    set note = '{note}',Name_web = '{Name_web}',sex = '{sex_web}',company = '{company_web}',
        company_type = '{company_type_web}',emp_type = '{emp_type_web}',tranning_type ='{tranning_type_web}' , update_date = '{updatetime}',note='{note}'
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