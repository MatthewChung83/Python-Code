def GetProxyIP(page):
    import requests
    from requests import Session
    from bs4 import BeautifulSoup
    import json
    from fake_useragent import UserAgent
    import csv
    
    ua = UserAgent()
    uar = ua.random
    
    url = f'https://www.xicidaili.com/wt/{page}'
    res = requests.get(url, headers={'User-Agent': uar})
    soup = BeautifulSoup(res.text,'html')
    
    ip_list = soup.find_all(id = 'ip_list')[0].find_all('tr')

    for i in range(1,len(ip_list)):
        
        global resp
        
        ip_info = ip_list[i].find_all('td')
        
        ip = ip_info[1].text
        port = ip_info[2].text
        country = ip_info[3].text
        level = ip_info[4].text
        segment = ip_info[5].text
        survival = ip_info[8].text
        valid_dt = ip_info[9].text
        
        proxy = {'http':'http://'+ ip + ':' + port,'https':'https://'+ ip + ':' + port}    
        url = 'http://easymap.land.moi.gov.tw/R02/Door_json_getCoordXY'
        req_session = Session()
        ua = UserAgent()
        uar = ua.random
        try:
            resp = req_session.post(url,proxies=proxy,timeout=20)
            
            if '200' in resp.status_code:
                data = (i,ip,port,country,level,segment,survival,valid_dt)
                file_path = r'C:\Users\cyche\Documents\Python Scripts\land_moi\proxy_list.csv'
                csvFileToWrite = open(file_path, 'a', encoding='utf-8-sig', newline='')
                csvDataToWrite = csv.writer(csvFileToWrite)
                csvDataToWrite.writerow(data)
                csvFileToWrite.close()
                print(data)
            else:
                print(ip,resp.status_code,'fail')
                
        except:
            print(ip,resp.status_code,'fail')

GetProxyIP(1)