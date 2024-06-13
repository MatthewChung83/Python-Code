from bs4 import BeautifulSoup
from requests import Session
import json

req_session = Session()
proxies = {
        'http': 'socks5://127.0.0.1:9150',
        'https': 'socks5://127.0.0.1:9150'
        }
url = 'https://api.ipify.org?format=json'

resp_before = req_session.get(url)
resp_after = req_session.get(url,proxies=proxies)

print(resp_before.text,resp_after.text)


