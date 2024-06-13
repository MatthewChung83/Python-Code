
##### 模擬瀏覽器行為 & 破解辨識碼

# 定位點擊html object的text，並於網站啟動反爬時，辨識驗證碼
def findelement_LINK_click(driver,value):
    
    from bs4 import BeautifulSoup
    from selenium.webdriver.common.by import By
    
    driver.find_element(By.LINK_TEXT, value).click()
    soup = BeautifulSoup(driver.page_source,'lxml')
    while anti_java(soup) == 'Y':
        save_captcha(driver)
        captcha()
        soup = captcha_in(driver)

# 定位點擊html object的名稱，並於網站啟動反爬時，辨識驗證碼
def findelement_CSS_click(driver,value):
    
    from bs4 import BeautifulSoup
    from selenium.webdriver.common.by import By
    
    driver.find_element(By.CSS_SELECTOR, value).click()
    soup = BeautifulSoup(driver.page_source,'lxml')
    while anti_java(soup) == 'Y':
        save_captcha(driver)
        captcha()
        soup = captcha_in(driver)

# 定位點擊html object的ID，並於網站啟動反爬時，辨識驗證碼
def findelement_ID_click(driver,value):
    
    from bs4 import BeautifulSoup
    from selenium.webdriver.common.by import By
    
    driver.find_element(By.ID, value).click()
    soup = BeautifulSoup(driver.page_source,'lxml')
    while anti_java(soup) == 'Y':
        save_captcha(driver)
        captcha()
        soup = captcha_in(driver)

# 定位傳送input 到 html object的ID，並於網站啟動反爬時，辨識驗證碼
def findelement_ID_sendkeys(driver,value,keys):
    
    from bs4 import BeautifulSoup
    from selenium.webdriver.common.by import By
    
    driver.find_element(By.ID, value).send_keys(keys)
    soup = BeautifulSoup(driver.page_source,'lxml')
    while anti_java(soup) == 'Y':
        save_captcha(driver)
        captcha()
        soup = captcha_in(driver)

# 定位點擊html object的xpath，並於網站啟動反爬時，辨識驗證碼   
def findelement_XPATH_href(driver,value):
    from bs4 import BeautifulSoup
    from selenium.webdriver.common.by import By
    
    driver.find_element_by_xpath('//a[@href="'+f'{value}'+'"]').click()
    soup = BeautifulSoup(driver.page_source,'lxml')
    while anti_java(soup) == 'Y':
        save_captcha(driver)
        captcha()
        soup = captcha_in(driver)

# 偵測反爬機制是否啟動，並於機制啟動後，辨識驗證碼
def defense_anti(driver):
    from bs4 import BeautifulSoup
    from selenium.webdriver.common.by import By
    
    soup = BeautifulSoup(driver.page_source,'lxml')
    while anti_java(soup) == 'Y':
        save_captcha(driver)
        captcha()
        soup = captcha_in(driver)

# 偵測反爬機制是否啟動
def anti_java(soup):
    from bs4 import BeautifulSoup
    if 'Please enable JavaScript to view the page content' in soup.text:
        result = 'Y'
    else:
        result = 'N'      
    return result

# 存取驗證碼圖片
def save_captcha(driver,print_path,captcha_path):
    from PIL import Image
    
    driver.save_screenshot(print_path)
    img = Image.open(print_path)
    new_img = img.crop((10, 30, 220, 100))
    new_img.save(captcha_path)

# 返回辨識之驗證碼文字    
def captcha():
    import ddddocr
    from PIL import Image
    
    ocr = ddddocr.DdddOcr()
    with open(captcha_path, 'rb') as f:
        img_bytes = f.read()
    
    res = ocr.classification(img_bytes)
    f.close()
    return res

# 輸入驗證碼文字
def captcha_in(driver):
    from bs4 import BeautifulSoup
    import pyautogui
    import time
    
    captext = captcha()
    pyautogui.click(243,233)
    time.sleep(1)
    pyautogui.typewrite(captext)
    time.sleep(1)
    pyautogui.press('enter')
    soup = BeautifulSoup(driver.page_source,'lxml')
    return soup






##### config 衍生

# 生成指定期間之日曆檔
def calendar(strdate,enddate):
    import datetime
    dateFormatter = "%Y/%m/%d"
    str_date = datetime.datetime.strptime(strdate, dateFormatter)
    end_date = datetime.datetime.strptime(enddate, dateFormatter)
    calendar = []
    
    for i in range((end_date - str_date).days + 1):
        dif_date = str_date + datetime.timedelta(days = i)
        year = str(int(dif_date.strftime("%Y")) -1911)
        month = str(dif_date.strftime("%m"))
        day = str(dif_date.strftime("%d"))
        calendar.append(year+'.'+month+'.'+day)
    return calendar





##### DB處理 相關

# 查詢資料庫是否已又相同法文link        
def src_obs(server,username,password,database,tb,columns,link):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select count ({columns})
    from [{tb}]
    where 
    [link] = '{link}'
    """    
    cursor.execute(script)
    obs = cursor.fetchall()
    cursor.close()
    conn.close()
    return list(obs[0])[0]

# 入庫
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

# 格式整理-法文主檔
def indata(doc):
    etl = []
    etl.append({
        'court':doc[0],
        'court_word':doc[1],
        'space':doc[2],
        'event':doc[3],        
        'ann_date':doc[4],
        'link':doc[5],
        'qtype':doc[6],
        's_parse':doc[7],
        'e_parse':doc[8],
        'main_sta':doc[9],
        'claimant':doc[10],
        'opposite':doc[11],
        'stakeholder':doc[12],
        'protester':doc[13],
        'plaintiff':doc[14],
        'defendant':doc[15],
        'dissenter':doc[16],
        'article':doc[17],
        'insertdate':doc[18],
        'filename':doc[19],	
        

    })
    return etl

# 格式整理-法文角色姓名主檔
def indetail(docs):
    etl = []
    for name in docs[2]:
        etl.append({
            "link" : docs[0],
            "character": docs[1],
            "name": name,
            "insertdate": docs[3],
            })
    return etl





##### 資料清洗處理 相關

# 資料整理-法文主檔切分
def parse(soup,content_tag1,content_tag2,splits,deletes):
    from bs4 import BeautifulSoup
    import re
    
    if soup.find(class_= content_tag1).text:
        content = soup.find(class_=content_tag1).text
    else:
        content = soup.find(class_=content_tag2).text
        
    if re.search(splits[0],content):
        s_parse = re.search(splits[0],content).span()[1]
    else:
        s_parse = 0
    
    if re.search('，',content):
        e_parse = re.search('，',content).span()[0]
    
    if e_parse == 99999:
        main_sta = content[0:1]
    else:
        main_sta = content[s_parse:e_parse]
    
    main_sta = main_sta[:3500]
    
    for de in deletes:
        main_sta = re.sub(de,'',main_sta)
    
    return content,s_parse,e_parse,main_sta

# 資料整理-角色切分
def roles_process(roles,content):
    import re
    
    role_seqs = []
    for role in roles:
        word = [_.group() for _ in re.finditer(role,content)]
        position = [_.span() for _ in re.finditer(role,content)]
        
        for w in range(len(word)):
            position[w]=list(position[w])
            position[w].append((role))
            
            if len(position[w]) > 0:
                role_seqs.append(position[w])
    role_seqs = sorted(role_seqs, key = lambda s: s[0])
    
    roles_ps = []
    for rs in range(len(role_seqs)):
        if rs == len(role_seqs) - 1:
            roles_p = [role_seqs[rs][2],role_seqs[rs][1],len(content)]
        else:
            roles_p = [role_seqs[rs][2],role_seqs[rs][1],role_seqs[rs+1][0]]
        roles_ps.append(roles_p)
    
    for rp in range(len(roles_ps)):
        roles_ps[rp].append(content[roles_ps[rp][1]:roles_ps[rp][2]])
    
    roles_content = []
    for ro in range(len(roles)):
        roles_content.append([roles[ro],''])
        for rp in range(len(roles_ps)):
            if roles_ps[rp][0] == roles[ro]:
                if len(re.sub(r'[\s]','',roles_ps[rp][3])) >= 2:
                    roles_content[ro][1] = roles_content[ro][1] + roles_ps[rp][3]
    return roles_content

# 資料整理-姓名清洗
def spname(delimis,text):
    import re
    
    for dl in delimis:
        text = re.sub(dl,',',text)
    texts = text.split(',')
    texts = list(filter(None,texts))
    names = []
    for t in texts:
        if len(t) <= 6 and len(t) >= 2:
            names.append(t)    
    names = list(set(names))
    return names