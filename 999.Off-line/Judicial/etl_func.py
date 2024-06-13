def src_obs(server,username,password,database,judicial_no,judicial_doc_no,publishing_date):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count(*)
    from [dbo].[judicial]
    where 
    judicial_no = '{judicial_no}' and [judicial_doc_no] = '{judicial_doc_no}' and publishing_date = '{publishing_date}'
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]

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

def Parse_Title(i,text,splits):
    str,end,k,l = 0,0,0,0
    str_w = ''
    while k < len(splits[i][1][0]):
        str = text.find(splits[i][1][0][k])
        if str > 0:
            str_w = splits[i][1][0][k]
            break
        k += 1
        
    while l < len(splits[i][1][1]):
        end = text.find(splits[i][1][1][l])
        if end > 0:
            break
        l += 1
    
    title_names = ''    
    if end > str:
        title_names = text[str + len(str_w) :end]    
    return title_names

def Parse_Subject(i,text,splits):
    str,end,m,n = 0,0,0,0
    str_w = ''
    while m < len(splits[i][2][0]):
        str = text.find(splits[i][2][0][m])
        if str > 0:
            str_w = splits[i][2][0][m]
            break
        m += 1
    
    while n < len(splits[i][2][1]):
        end = text.find(splits[i][2][1][n])    
        if end > 0:
            break
        n += 1
        
    subject_names = ''
    if end > str:
        subject_names = text[str + len(str_w) :end]    
    return subject_names

def Parse_anncm(i,text,splits):
    str,end,o,p = 0,0,0,0
    str_w = ''
    while o < len(splits[i][3][0]):
        str = text.find(splits[i][3][0][o])
        if str > 0:
            str_w = splits[i][3][0][o]
            break
        o += 1
    
    while p < len(splits[i][3][1]):
        end = text.find(splits[i][3][1][p])    
        if end > 0:
            break
        p += 1
        
    anncm_names = ''
    if end > str:
        anncm_names = text[str + len(str_w) :end]        
    return anncm_names

def delws_re(text,delws,delimis):
    import re

    for delw in delws:
        text = re.sub(delw,',',text)

    for d in range(len(delimis)):
        text = text.replace(delimis[d],',')

    text_list = list(set(text.split(',')))
    name_list = [i for i in text_list if i != '' and len(i) > 1 and len(i) < 15]
    text = ','.join(name_list)
    return text,name_list

def indata(doc):
    etl = []
    etl.append({
        'keyword':doc[0],
        'court':doc[1],
        'judicial_no':doc[2],
        'recipient':doc[3],
        'area':doc[4],
        'doc_type':doc[5],
        'publishing_date':doc[6],
        'judicial_type':doc[7],
        'judicial_title':doc[8],
        'judicial_doc_date':doc[9],
        'judicial_doc_no':doc[10],
        'judicial_doc_subject':doc[11],
        'judicial_doc_basis':doc[12],
        'judicial_doc_anncm':doc[13],
        'judicial_doc_pdate':doc[14],
        'judicial_doc_udate':doc[15],
        'judicial_doc_unit':doc[16],
        'judicial_content':doc[17],
        'insertdate':doc[18],
        'names':doc[19],
        'file_save':doc[20],
    })
    return etl

def indetail(doc):
    etl = []
    for debtor in doc[7]:
        etl.append({
            "judicial_no" : doc[0],
            "recipient": doc[1],
            "publishing_date": doc[2],
            "judicial_doc_date": doc[3],
            "judicial_doc_no": doc[4],
            "judicial_doc_subject": doc[5],
            "debtors": doc[6],
            "debtor": debtor,
            "insertdate": doc[8],
            })
    return etl