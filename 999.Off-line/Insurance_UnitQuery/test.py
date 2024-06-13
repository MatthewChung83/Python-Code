import requests
from operator import itemgetter, attrgetter
import time
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
import re

from config import *
from etl_func import *

text = '王秸宸,王玉輝,之一百一十一年二月十五日,王英州,王碧霞,第123人,王玉輝,之111年,等二十五人'

delws = ['第[1234567890]+人','[1234567890]+人','之[0123456789]+年',
         '等[一二三四五六七八九十]+人','等[一二三四五六七八九十]+','[一二三四五六七八九十]+人','上[一二三四五六七八九十]+' ,
         '之[一二三四五六七八九十百]+年[一二三四五六七八九十]+月[一二三四五六七八九十]+日',
         '[a-zA-Z0-9。:*：，\s]']

for delw in delws:
    text = re.sub(delw,',',text)


print(text)