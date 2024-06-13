# 測試名單量


# 測試資料庫資訊
db = {
    'server': 'vnt07.ucs.com',
    'database': 'uis',
    'username': 'pyuser',
    'password': 'Ucredit7607',
    
    # 測試名單
    'totb1':'base_case',
    # 投保資料主檔
    'totb2':'Insurance_UnitQuery',
}

# 呼叫之API
wbinfo = {
    'url':'https://www.nhi.gov.tw/OnlineQuery/Insurance_UnitQuery.aspx',
}

# user-agent

def parse(url,id):
    import requests
    from lxml import etree
    from bs4 import BeautifulSoup
    
    headers = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36'
    }
    
    s = requests.session()
    r = s.post(url,headers=headers)
    res = etree.HTML(r.text)
    
    data = {'__VIEWSTATE': res.xpath('//input[@id="__VIEWSTATE"]/@value')[0],
            'ctl00$ContentPlaceHolder1$tbxID': id,
            'ctl00$ContentPlaceHolder1$btnIDSend': '查詢',
            '__VIEWSTATEGENERATOR': res.xpath('//input[@id="__VIEWSTATEGENERATOR"]/@value')[0],
            '__SCROLLPOSITIONX': '0',
            '__SCROLLPOSITIONY': '0',
            '__EVENTTARGET': '',
            '__EVENTARGUMENT': '',
            '__VIEWSTATEENCRYPTED': '',
            '__EVENTVALIDATION': res.xpath('//input[@id="__EVENTVALIDATION"]/@value')[0],
            }
    resp = s.post(url,headers=headers,data=data)
    soup = BeautifulSoup(resp.text,"lxml")
    
    return soup

