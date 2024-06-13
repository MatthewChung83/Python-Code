
from bs4 import BeautifulSoup
from requests import Session
import json
from fake_useragent import UserAgent


CoordXY_url = 'http://easymap.land.moi.gov.tw/R02/Door_json_getCoordXY'
proxies = {
    'http': 'socks5h://127.0.0.1:9051',
    'https': 'socks5h://127.0.0.1:9051'}

req_session = Session()
#resp = req_session.post(CoordXY_url)
resp = req_session.post(CoordXY_url, proxies=proxies)
print(resp.status_code)
