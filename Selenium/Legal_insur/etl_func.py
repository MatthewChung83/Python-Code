# -*- coding: utf-8 -*-
"""
Created on Mon May  9 18:38:29 2022

@author: admin
"""




def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num
def src_obs(server,username,password,database,fromtb,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*) from INS_Legal_Insurtech where STATUS = 'N'
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]



def dbfrom(server,username,password,database,fromtb,totb,today):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    
    script = f"""
    with data as(
    SELECT
        [UUID]
        ,[DataDate]
        ,[LDCI]
        ,[CaseI]
        ,[DebtorName]
        ,[Legal_Num]
        ,[Legal_Court]
        ,STUFF((SELECT ',' + [DebtorID]
            FROM [INS_Legal_Insurtech]
            WHERE CaseI = s.CaseI and DataDt = '{today}'
            FOR XML PATH('')), 1, 1, '') AS [DebtorID]
        ,[Order_Num]
        ,[Transfer_Bank]
        ,[Transfer_Account]
        ,[Transfer_Fee]
        ,[Payment_Type]
        ,[Product_List]
        ,[Payment_Deadline]
        ,[Account_Name]
        ,[Legal_Type]
        ,[Insur_Type]
        ,[Notes]
        ,[ApplyName]
        ,[STATUS]
        ,[DataDt]
        ,seq = ROW_NUMBER()over(partition by s.casei order by s.[DebtorID])
    FROM [UCS_ReportDB].[dbo].[INS_Legal_Insurtech] s
    where STATUS = 'N' and DataDt = '{today}'
    )
    select * from data where seq = 1
    ORDER BY Casei
	offset 0 row fetch next 1 rows only
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    
    cursor.close()
    conn.close()
    return c_src


def update(server,username,password,database,totb,Notes,Casei,Order_Num,Account_Name,Payment_Type,Product_List,Transfer_Bank,Transfer_Account
           ,Payment_Deadline,Transfer_Fee,Legal_Type,Status,today):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database, charset = 'CP936')
    cursor = conn.cursor()
        
    script = f"""
    update [{totb}]
    set Notes = '{Notes}',Order_Num = '{Order_Num}',Account_Name='{Account_Name}',Payment_Type='{Payment_Type}',Product_List='{Product_List}'
        ,Transfer_Bank='{Transfer_Bank}',Transfer_Account='{Transfer_Account}',Payment_Deadline='{Payment_Deadline}',Transfer_Fee='{Transfer_Fee}'
        ,Legal_Type='{Legal_Type}',Status='{Status}'
    where casei = '{Casei}' and DataDt = '{today}'
    
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
   
