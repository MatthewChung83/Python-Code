# -*- coding: utf-8 -*-

from bs4 import BeautifulSoup
import requests
import ddddocr
import time
import re

from etl_func import *
from config import *
from dict import *

# 宣告資料庫、網站、迭代...等參數
server,database,username,password,fromtb,totb = db['server'],db['database'],db['username'],db['password'],db['fromtb'],db['totb']
url,captchaImg = wbinfo['url'],wbinfo['captchaImg']
imgp = pics['imgp']
q = 0

def dbfrom(server,username,password,database,fromtb,totb):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    select 
    personi, ID, Name COLLATE Chinese_PRC_CI_AS, C, M
    from [dbo].[{fromtb}]
    where type = 'TEST' and personi not in (select personi from [{totb}] where [status] in ('Y','N'))
    order by personi
    """
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

print(dbfrom(server,username,password,database,fromtb,totb)[0][2].encode('latin1').decode('gbk'))