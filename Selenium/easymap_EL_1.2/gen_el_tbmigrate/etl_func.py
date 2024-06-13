def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num

def dbfromel07(server,username,password,database,fromtb,Machine):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select a.entitytype,success_jobs 
    from 
    (
    select entitytype,sum(case when status = 'Done' then 1 else 0 end) as success_jobs 
    from [dbo].[EL_Entity_Managerment] group by entitytype
    ) as a
    where success_jobs = 4
    """
    cursor.execute(script)
    recovery_index = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return recovery_index

def el04migrate(server,username,password,database,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    INSERT INTO [dbo].[{totb}]
    SELECT * FROM [dbo].[{totb}_TMP_{entitytype}] as src
    WHERE NOT EXISTS (SELECT el_number FROM [dbo].[{totb}] as ta WHERE ta.el_number = src.el_number)
    """    
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    return script

def el05migrate(server,username,password,database,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    INSERT INTO [dbo].[{totb}]
    SELECT * FROM [dbo].[{totb}_TMP_{entitytype}] as src
    WHERE NOT EXISTS (SELECT el_number FROM [dbo].[{totb}] as ta WHERE ta.el_number = src.el_number)
    """    
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    return script