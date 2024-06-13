def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num

def src_obs(server,username,password,database,fromtb,totb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*)
    from [dbo].[{fromtb}]
    where type_query in ('{entitytype}') and (status not in ('Y','N','X') or status is null)
    """                                                     #20211004 Lillian+status=X>>姓名難字無法查詢
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
    where psid = {psid} and pid = '{pid}' and ci = {ci} and register_no = '{register_no}' and change_no = '{change_no}' and type_query = '{entitytype}'
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]

def dbfrom(server,username,password,database,fromtb,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select psid,pid,cname,ci,type
    from [dbo].[{fromtb}] as a
    where type_query in ('{entitytype}') and (status not in ('Y','N','X') or status is null)
    order by psid,pid,ci
    offset 0 row fetch next 1 rows only
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src
    
def toSQL(docs, totb, server, database, username, password):
    import pyodbc
    conn_cmd = 'DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password
    #conn_cmd = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password
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
            
def mp_etl(doc):
    mp_tmp_etl = []
    mp_tmp_etl.append({
        "psid":doc[0],
        "pid":doc[1],
        "ci":doc[2],
        "cname":doc[3],        # '債務人姓名',
        "item":doc[4],           # '項次',
        "status":doc[5],           # '案件狀態',
        "register_organ":doc[6],        # '登記機關',
        "class":doc[7],        # '案件類別',
        "register_no":doc[8],        # '登記編號',
        "register_date":doc[9],        # '登記核准日期',
        "change_no":doc[10],        # '變更文號',
        "change_date":doc[11],        # '變更核准日期',
        "logout_no":doc[12],        # '註銷文號',
        "logout_date":doc[13],        # '註銷日期',
        "agent_cname":doc[14],        # '債務代理人名稱',
        "agent_pid":doc[15],        # '債務代理人統編',
        "creditor_name":doc[16],        # '債權人名稱',
        "creditor_id":doc[17],        # '債權人統編',
        "creditor_agent_name":doc[18],        # '債權代理人名稱',
        "creditor_agent_id":doc[19],        # '債權代理人統編',
        "contract_start_date":doc[20],      # '契約啟始日期',
        "contract_end_date":doc[21],      # '契約終止日期',
        "subject_owner_name":doc[22],      # '標的物所有人名稱',
        "estate_amt":doc[23],      # '擔保債權金額',
        "subject_owner_id":doc[24],      # '標的物所有人統編',
        "estate_items":doc[25],      # '動產明細項數',
        "subject_address":doc[26],      # '標的物所在地',
        "limitation_flg":doc[27],      # '是否最高限額',
        "float_flg":doc[28],      # '是否為浮動擔保',
        "subject_species":doc[29],     # '標的物種類'
        "data_date":doc[30],  #  資料更新時間
        "printscreen":doc[31],
        "type_query":doc[32],
    })
    return mp_tmp_etl

def updateSQL(server,username,password,database,fromtb,psid,pid,ci,status,entitytype):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    update PropertySecuredtb set status = '{status}'
    where psid={psid} and pid = '{pid}' and ci = {ci} and type_query = '{entitytype}';
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    return script