def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num

def src_obs(server,username,password,database,totb1,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select (select count(*)
    from [dbo].[{totb1}]
    where (status is null)) --entitytype in ('{entitytype}') and (status is null))
    +
    (select count(*) from [{totb1}] 
    where (select DATEDIFF(mm, updatetime, getdate()))> = 3 and status <> 'Y')
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]

def target_obs(server,username,password,database,totb,psid,pid,ci,register_no,change_no,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*)
    from [dbo].[{totb}]
    where psid = {psid} and pid = '{pid}' and ci = {ci} and register_no = '{register_no}' and change_no = '{change_no}' --and type_query = '{entitytype}'
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
    into #test
    from [dbo].[{totb1}] as a
    where (status is null) --entitytype in ('{entitytype}') and (status is null)
    
    insert into #test
    select  *
    from [{totb1}]
    where (select DATEDIFF(mm, updatetime, getdate()))> = 3 and status <> 'Y'
    
    select * from #test
    order by entitytype desc
    offset 0 row fetch next 1 rows only
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def updateSQL(server,username,password,database,totb1,entitytype,estateAgent_inc,EAvaliddate,status,updatetime,ID,name):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    update [dbo].[{totb1}]
    set estateAgent_inc = '{estateAgent_inc}',EAvaliddate = '{EAvaliddate}', status = '{status}', updatetime = '{updatetime}'
    where ID = '{ID}' -- entitytype in ('{entitytype}') and ID = '{ID}' --and name = '{name}'
    """
    
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
def exit_obs(server,username,password,database,totb1):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count (distinct ID)
    from [{totb1}]
    where  updatetime > = convert(varchar(10),getdate(),111)
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]