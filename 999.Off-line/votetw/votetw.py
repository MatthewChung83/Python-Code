# -*- coding: utf-8 -*-
"""
Created on Thu Jan 20 22:42:02 2022

@author: admin
"""


import datetime
import time
import requests
import sys
import re

from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from requests import Session



from config import *
from etl_func import *
from dict import *

server,database,username,password,totb= db['server'],db['database'],db['username'],db['password'],db['totb']
url = wbinfo['url']

obs = src_obs(server,username,password,database,totb)
print(src_obs(server, username, password, database,totb))
for i in foo(-1,obs-1):
    src = dbfrom(server,username,password,database,totb)[0]
    #rowid
    rowid = str(src[0])
    vote = src[1]
    #選舉區
    vote_city = src[2]
    print(vote_city)
    vote_town = src[3].replace('選舉區','選區')
    print(vote_town)
    vote_village = src[4].replace('選舉區','選區')
    print(vote_village)
    #print(vote_area.split('\n')[1])
    #調閱姓名
    name = src[6]
    #排除英文字
    name = re.sub('[a-zA-Z]','',name)
    #排除空格及特定符號
    name = name.replace(' ','').replace('‧','').replace('．','')
    
    new_url = url + name
    print(new_url)
    resp = requests.get(new_url)
    soup = BeautifulSoup(resp.text,"lxml")
    href = soup.findAll('a')
    for i in range(len(href)) :
        url_href = soup.findAll('a')[i].get('href')
        name_href = str(soup.findAll('a')[i].text)
        name_href = re.sub('[a-zA-Z]','',name_href).replace(' ','').replace('‧','').replace('．','')
        if name_href == name:
            print(url_href)
            print(name_href)
            req = requests.get(url_href)
            soup_req = BeautifulSoup(req.text,"lxml")
            for i in range(len(soup_req.findAll(class_='col-xs-12')[1].findAll('tr'))):
                if i == 0:
                    pass
                else:                    
                    area_req = soup_req.findAll(class_='col-xs-12')[1].findAll('tr')[i].findAll('td')[1].text.replace('01','1').replace('02','2').replace('03','3').replace('04','4').replace('05','5').replace('06','6').replace('07','7').replace('08','8').replace('09','9')
                    birthday = soup_req.findAll(class_='col-xs-12')[1].findAll('tr')[i].findAll('td')[2].text
                    vote_n = soup_req.findAll(class_='col-xs-12')[1].findAll('tr')[i].findAll('td')[7].text
                    print(vote_village,len(vote_village))
                    print(area_req,len(area_req))
                    print(birthday)
                    print(vote_n)
                    
                    if vote_city.replace(' ','') in area_req.replace(' ','') and vote_town.replace(' ','') in area_req.replace(' ','') and vote_village.replace(' ','') in area_req.replace(' ','') and '月' in birthday :
                        area_check = '1_縣市、行政區域、選區完整比對'
                        update(server,username,password,database,totb,birthday,area_req,vote_n,rowid,name_href,area_check)
                        break
                        
                    elif vote_town.replace(' ','') in area_req.replace(' ','') and len(vote_town)> 0  and '月' in birthday:
                        #print('B')
                        area_check = '2_行政區域比對'
                        update(server,username,password,database,totb,birthday,area_req,vote_n,rowid,name_href,area_check)
                        break
                    elif vote_city.replace(' ','') in area_req.replace(' ','')  and '月' in birthday:
                        #print('C')
                        area_check = '3_縣市比對'
                        update(server,username,password,database,totb,birthday,area_req,vote_n,rowid,name_href,area_check)
                        break
                    
                    #elif vote_village.replace(' ','')  in area_req[0:3].replace(' ','')  and '月' in birthday:
                    #    area_check = '1'
                    #    update(server,username,password,database,totb,birthday,area_req,vote_n,rowid,name_href,area_check)
                    #    continue
                    #elif vote_town in area_req[0:3] and '月' in birthday:
                    #    area_check = '2'
                    #    update(server,username,password,database,totb,birthday,area_req,vote_n,rowid,name_href,area_check)
                    #    continue
                    #elif vote_city in area_req[0:3] and '月' in birthday:
                    #    area_check = '3'
                    #    update(server,username,password,database,totb,birthday,area_req,vote_n,rowid,name_href,area_check)
                    #    continue
            continue
        else :
            birthday = ''
            area_req = ''
            vote_n = ''
            name_href = ''
            area_check = ''
            update(server,username,password,database,totb,birthday,area_req,vote_n,rowid,name_href,area_check)
            pass
        
    
    
        

        


