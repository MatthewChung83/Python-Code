### ??��??建�?? ###
def gettowninfo(url,cityCode,cityName,doorPlateType):
    from requests import Session
    from fake_useragent import UserAgent
    from bs4 import BeautifulSoup
    
    uar = UserAgent().random
    req_session = Session()
    captchacheck_url = 'https://easymap.land.moi.gov.tw/CaptchaCheck_json_showCheck'
    resp = req_session.post(captchacheck_url)
    cookie = '; '.join( [ '='.join(i) for i in resp.cookies.items()] )
    
    headers={
        #'Cookie': cookie,
        'Host': 'easymap.land.moi.gov.tw',
        'Origin': 'https://easymap.land.moi.gov.tw',
        'Referer': 'https://easymap.land.moi.gov.tw/Index',
        'User-Agent': uar,
    }
    
    set_token_url = 'https://easymap.land.moi.gov.tw/pages/setToken.jsp'
    token_resp = req_session.post(set_token_url, headers=headers)
    token = BeautifulSoup( token_resp.text, 'lxml').select('input[name=token]')[0]['value']

    data = {
        'cityCode':cityCode,
        'cityName':cityName,
        'doorPlateType':doorPlateType,
        'struts.token.name':'token',
        'token': token,
    }

    resp=req_session.post(url, headers=headers, data=data)
    return resp.text

def getdoorinfo(url,cityCode,towncode,road,doorPlateType,lane,alley,area,cityName,no):
    from requests import Session
    from fake_useragent import UserAgent
    from bs4 import BeautifulSoup
    
    uar = UserAgent().random    
    req_session = Session()
    captchacheck_url = 'https://easymap.land.moi.gov.tw/P02/CaptchaCheck_json_showCheck'
    resp = req_session.post(captchacheck_url)
    cookie = '; '.join( [ '='.join(i) for i in resp.cookies.items()] )
    
    headers={
        #'Cookie': cookie,
        'Host': 'easymap.land.moi.gov.tw',
        'Origin': 'https://easymap.land.moi.gov.tw',
        'Referer': 'https://easymap.land.moi.gov.tw/P02/Index',
        'User-Agent': uar,
    }
    
    set_token_url = 'https://easymap.land.moi.gov.tw/P02/pages/setToken.jsp'
    token_resp = req_session.post(set_token_url, headers=headers)
    token = BeautifulSoup( token_resp.text, 'lxml').select('input[name=token]')[0]['value']

    data = {
        'city': cityCode,
        'area': towncode,
        'road': road,
        'doorPlate': road,
        'doorPlateType': doorPlateType,
        'lane': lane,
        'alley': alley,
        'townName': area,
        'cityName': cityName,
        'no': no,
        'struts.token.name': 'token',
        'token': token,
    }

    resp=req_session.post(url, headers=headers, data=data)
    return resp.text

def getbuildinfo(url,office,sectno,buildno):
    from requests import Session
    from fake_useragent import UserAgent
    from bs4 import BeautifulSoup
    
    uar = UserAgent().random    
    req_session = Session()
    captchacheck_url = 'https://easymap.land.moi.gov.tw/P02/CaptchaCheck_json_showCheck'
    resp = req_session.post(captchacheck_url)
    cookie = '; '.join( [ '='.join(i) for i in resp.cookies.items()] )
    
    headers={
        #'Cookie': cookie,
        'Host': 'easymap.land.moi.gov.tw',
        'Origin': 'https://easymap.land.moi.gov.tw',
        'Referer': 'https://easymap.land.moi.gov.tw/Index',
        'User-Agent': uar,
    }
    
    set_token_url = 'https://easymap.land.moi.gov.tw/P02/pages/setToken.jsp'
    token_resp = req_session.post(set_token_url, headers=headers)
    token = BeautifulSoup( token_resp.text, 'lxml').select('input[name=token]')[0]['value']

    data = {
        'office': office,
        'sectNo': sectno,
        'buildingNo': buildno,
        'struts.token.name': 'token',
        'token': token,
    }

    resp=req_session.post(url, headers=headers, data=data)
    return resp.text

### ??��????��?? ###
def GetCoordXY(url,cityName,road,lane,alley,no,area):
    from bs4 import BeautifulSoup
    from requests import Session
    from fake_useragent import UserAgent

    uar = UserAgent().random
    req_session = Session()
    captchacheck_url = 'https://easymap.land.moi.gov.tw/R02/CaptchaCheck_json_showCheck'
    resp = req_session.post(captchacheck_url)
    cookie = '; '.join( [ '='.join(i) for i in resp.cookies.items()] )    

    headers={
        #'Cookie': cookie,
        'Host': 'easymap.land.moi.gov.tw',
        'Origin': 'https://easymap.land.moi.gov.tw',
        'Referer': 'https://easymap.land.moi.gov.tw/R02/Index',
        'User-Agent': uar,
    }
    
    set_token_url = 'https://easymap.land.moi.gov.tw/R02/pages/setToken.jsp'
    token_resp = req_session.post(set_token_url, headers=headers)
    token = BeautifulSoup(token_resp.text, 'lxml').select('input[name=token]')[0]['value']
    rlan = road+lane+alley+no
    
    data = {
        'cityName': cityName,
        'doorPlate': rlan,
        'townName': area,
        'struts.token.name': 'token',
        'token': token,
    }
    resp=req_session.post(url, headers=headers, data=data,)
    CoordXY = BeautifulSoup(resp.text,'html.parser')
    return CoordXY

def GetDoorInfoByXY(url,cityCode,CoordX,CoordY,road,lane,alley,no):
    from bs4 import BeautifulSoup
    from requests import Session
    from fake_useragent import UserAgent

    uar = UserAgent().random
    req_session = Session()
    captchacheck_url = 'https://easymap.land.moi.gov.tw/P02/CaptchaCheck_json_showCheck'
    resp = req_session.post(captchacheck_url)
    cookie = '; '.join( [ '='.join(i) for i in resp.cookies.items()] )    

    headers={
        #'Cookie': cookie,
        'Host': 'easymap.land.moi.gov.tw',
        'Origin': 'https://easymap.land.moi.gov.tw',
        'Referer': 'https://easymap.land.moi.gov.tw/P02/Index',
        'User-Agent': uar,
    }
    
    set_token_url = 'https://easymap.land.moi.gov.tw/P02/pages/setToken.jsp'
    token_resp = req_session.post(set_token_url, headers=headers)
    token = BeautifulSoup(token_resp.text, 'lxml').select('input[name=token]')[0]['value']
    rlan = road+lane+alley+no
    
    data = {
        'city': cityCode,
        'coordX': CoordX,
        'coordY': CoordY,
        'struts.token.name': 'token',
        'token': token,
    }

    resp=req_session.post(url, headers=headers, data=data)
    DoorInfoByXY = BeautifulSoup(resp.text,'html.parser')
    return DoorInfoByXY

def GetLandDesc(url,cityCode,towncode,office,sectno,landno):

    from bs4 import BeautifulSoup
    from requests import Session
    from fake_useragent import UserAgent

    uar = UserAgent().random
    req_session = Session()
    captchacheck_url = 'https://easymap.land.moi.gov.tw/R02/CaptchaCheck_json_showCheck'
    resp = req_session.post(captchacheck_url)
    cookie = '; '.join( [ '='.join(i) for i in resp.cookies.items()] )    

    headers={
        #'Cookie': cookie,
        'Host': 'easymap.land.moi.gov.tw',
        'Origin': 'https://easymap.land.moi.gov.tw',
        'Referer': 'https://easymap.land.moi.gov.tw/R02/Index',
        'User-Agent': uar,
    }
    
    set_token_url = 'https://easymap.land.moi.gov.tw/R02/pages/setToken.jsp'
    token_resp = req_session.post(set_token_url, headers=headers)
    token = BeautifulSoup(token_resp.text, 'lxml').select('input[name=token]')[0]['value']
    
    data = {
        'cityCode': cityCode,
        'townCode': towncode,
        'office': office,
        'sectNo': sectno,
        'landNo': landno,
        'struts.token.name': 'token',
        'token': token,
    }

    resp=req_session.post(url, headers=headers, data=data)
    return resp.text