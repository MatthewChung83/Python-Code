from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from requests.exceptions import Timeout
from requests import exceptions
from selenium import webdriver
from bs4 import BeautifulSoup

from PIL import Image
import urllib
import requests
import time
import re
import io

# import import_ipynb
from etl_func import *
from config import *

server,database,username,password,totb = db['server'],db['database'],db['username'],db['password'],db['totb']
truncate(server,username,password,database)