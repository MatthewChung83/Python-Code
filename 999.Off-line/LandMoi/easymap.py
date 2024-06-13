from bs4 import BeautifulSoup
from requests import Session

def main(cityCode, townCode, office, sectNo, landNo):
    req_session = Session()
    captchacheck_url = 'http://easymap.land.moi.gov.tw/R02/CaptchaCheck_json_showCheck'
    resp = req_session.post(captchacheck_url)
    cookie = '; '.join( [ '='.join(i) for i in resp.cookies.items()] )

    headers={
        'Cookie': cookie,
        'Host': 'easymap.land.moi.gov.tw',
        'Origin': 'http://easymap.land.moi.gov.tw',
        'Referer': 'http://easymap.land.moi.gov.tw/R02/Index',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36',
    }

    # get token
    set_token_url = 'http://easymap.land.moi.gov.tw/R02/pages/setToken.jsp'
    token_resp = req_session.post(set_token_url, headers=headers)
    token = BeautifulSoup( token_resp.text, 'lxml').select('input[name=token]')[0]['value']

    data = {
        'cityCode': cityCode,
        'townCode': townCode,
        'office': office,
        'sectNo': sectNo,
        'landNo': landNo,
        'struts.token.name': 'token',
        'token': token,
    }
    url='http://easymap.land.moi.gov.tw/R02/LandDesc_ajax_detail'
    resp=req_session.post(url, headers=headers, data=data)

    result = {}
    doc=BeautifulSoup(resp.text,'lxml')
    for row in doc.select('tr'):
        row_name = row.th
        if row_name:
            result[row_name.text]=row.td.text
    return result

if __name__=="__main__":
    cityCode='O'
    townCode='O01'
    office='OA'
    sectNo='0032'
    landNo='145'
    print( main(cityCode, townCode, office, sectNo, landNo) )
