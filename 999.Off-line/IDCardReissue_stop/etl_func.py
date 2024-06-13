def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num

def src_obs(server,username,password,database,totb1,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*)
    from [dbo].[{totb1}]
    where entitytype in ('{entitytype}') and (status is null)
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]

def dbfrom(server,username,password,database,totb1,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select *
    from [dbo].[{totb1}] as a
    where entitytype in ('{entitytype}') and (status is null)
    order by id
    offset 0 row fetch next 1 rows only
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def updateSQL(server,username,password,database,totb1,entitytype,status,updatetime,ID,name,note):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    update [dbo].[{totb1}]
    set note = '{note}',status = '{status}',updatetime = '{updatetime}'
    where entitytype in ('{entitytype}') and ID = '{ID}' and name = '{name}'
    """
    
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()

