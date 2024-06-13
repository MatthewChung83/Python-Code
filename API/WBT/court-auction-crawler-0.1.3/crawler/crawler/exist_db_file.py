# add initial & incremental src, add by Chris 2021/3/1
def exist_db():
    import pymssql
    conn = pymssql.connect(server='10.10.0.94', user='CLUSER', password='Ucredit7607', database='CL_Daily')
    cursor = conn.cursor()
    script = f"""select distinct CONCAT(court,'_',number,'_',REPLACE(REPLACE(date,' ','_'),'/',''),'.pdf') from wbt_court_auction_tb"""
    cursor.execute(script)
    qry = cursor.fetchall()
    cursor.close()
    conn.close()
          
    exist_pdf = []        
    for i in range(len(qry)):
        pdf = qry[i][0]
        exist_pdf.append(pdf)
    return exist_pdf