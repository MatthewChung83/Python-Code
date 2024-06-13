def foo(num,obs):
    while num < obs:
        num = num + 1 
        yield num

def toSQL(docs, totb, server, database, username, password):
    import pyodbc
    # conn_cmd = 'DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password
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

def indata(doc):
    etl = []
    etl.append({
        'city_id':doc[0],
        'city':doc[1],
        'area_id':doc[2],
        'area':doc[3],
        'parcel_section_id':doc[4],        
        'parcel_section':doc[5],
        'isword':doc[6],
        'updatetime':doc[7],
    })
    return etl
    
def capcha_resp(driver,respimg,q,acc_name,acc_pwd):
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.wait import WebDriverWait
    from selenium.common.exceptions import TimeoutException
    from selenium.webdriver.common.by import By
    from bs4 import BeautifulSoup
    from PIL import Image
    import requests
    import ddddocr
    import time
    import io
    
    while q < 100:
        
        driver.find_element(By.ID,'aa-uid').clear()
        driver.find_element(By.ID,'aa-passwd').clear()
        
        driver.find_element(By.ID,'aa-uid').send_keys(acc_name)
        driver.find_element(By.ID,'aa-passwd').send_keys(acc_pwd)
        
        img_ele = driver.find_element_by_id('AAAIden1')
        img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
        img.save(respimg)
        
        ocr = ddddocr.DdddOcr()
        with open(respimg, 'rb') as f:
            img_bytes = f.read()
            res = ocr.classification(img_bytes)

        driver.find_element(By.NAME,'aa-captchaID').send_keys(res)
        driver.find_element(By.ID,'submit_hn').click()
        
        try:
            WebDriverWait(driver, 3).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            time.sleep(2)
            alert.accept()
            time.sleep(2)
            msg = 'alert accepted'
        
        except TimeoutException:
            msg = 'login in'
            return msg
            break
        q += 1

def is_all_chinese(strs):
    for _char in strs:
        if not '\u4e00' <= _char <= '\u9fa5':
            return False
    return True

def truncate(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    truncate table [uis].[dbo].[land_parcel_section_tmp];
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()
    
def overwrite(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    truncate table [uis].[dbo].[land_parcel_section]
    INSERT INTO [uis].[dbo].[land_parcel_section] SELECT * FROM [uis].[dbo].[land_parcel_section_tmp]
    """
    cursor.execute(script)
    conn.commit()
    cursor.close()
    conn.close()

def dbtest(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    truncate table [uis].[dbo].[land_parcel_section_tmp]
    """
    cursor.execute(script)
    # c_src = cursor.fetchall()
    conn.commit()
    cursor.close()
    conn.close()
    # return c_src