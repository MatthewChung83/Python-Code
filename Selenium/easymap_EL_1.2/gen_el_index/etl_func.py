def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num

def dbfromel02(server,username,password,database,fromtb,Machine):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select [EntityType], [Machine], [EntityPhase], [EntityPath], [EntityObs], [validtrans], [Status], [actiondttm], [Statusdttm]
    from [{fromtb}] where Machine = '{Machine}' and status <> 'Done' and EntityPhase = 'el02_main.py'
    order by EntityType
    """
    cursor.execute(script)
    batch_index = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return batch_index

def dbfromel03(server,username,password,database,fromtb,Machine):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    select [EntityType], [Machine], [EntityPhase], [EntityPath], [EntityObs], [validtrans], [Status], [actiondttm], [Statusdttm]
    from [{fromtb}] where Machine = '{Machine}' and status <> 'Done' and EntityPhase = 'el03_main.py'
    order by EntityType
    """
    cursor.execute(script)
    batch_index = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return batch_index

