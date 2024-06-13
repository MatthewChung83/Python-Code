#!/usr/bin/env python
# -*- coding: utf-8 -*-

import regex as re
import traceback
import requests,pandas
from datetime import datetime
from bs4 import BeautifulSoup

# main function
def parseBulletin(tfasc,mode='prod'):
    if mode=='dev': tfasc = tfasc[:10]
    retry=3
    result = []
    df = pandas.DataFrame([_ for _ in tfasc]).fillna('')
    df_g = df.groupby(['document','金服案號'])
    all_len = len(df_g)
    print( 'total: %d'%(all_len) )
    _index = 1
    for document,number in df_g.groups.keys():
        if (_index%100)==0:
            print( "%d/%d"%(_index,all_len) )
        for i in range(retry):
            try:
                if len(document)== 0 or len(number) == 0:
                    break
                else:
                    result += _parse_bulletin(df,df_g,document,number)
                    break
            except:
                traceback.print_exc()
            print(result)
        _index+=1
    return result

def _parse_bulletin(df,df_g,document,number):
    print(document,number)
    result = []
    if len(document)== 0 or len(number) == 0:
        pass
    else:
        land_tb,build_tb,court,court_number,owner_dict,owners_org = getBulletin(document,number)
        

    unit_index=0
    unit_list = get_build_number(build_tb)
    for _,doc in df_g.get_group( (document,number) ).iterrows():
        doc = address_split(doc)
        doc['court'],doc['court_number'],doc['owners_org'] = court,court_number,owners_org
        doc['owner'] = owner_dict.get(doc['標別'],'')
        doc['unit']=''
        if doc['remark']=='建物':
            doc['unit'] = unit_list[unit_index]
            unit_index+=1
        result.append(doc.to_dict())
    return result

def parseSection(mode='prod'):
    tfasc=[]
    retry=3
    section = getSection()
    if mode=='dev': section = section[:1]
    for data in section:
        for i in range(retry):
            try:
                tfasc+=getBuzBidTime(data)
                break
            except:
                traceback.print_exc()
    tfasc = etl(tfasc)
    return tfasc
###
def num_transformer(num):
    _num_dict = {'１':'1', '２':'2', '３':'3', '４':'4', '５':'5', '６':'6', '７':'7', '８':'8', '９':'9', '０':'0'}
    for k,v in _num_dict.items(): num = num.replace(k,v)
    return num

def getSection():
    url = 'https://www.tfasc.com.tw/Product/BuzBidTime/ReadData'
    data = {
        "sort": "",
        "page": "1",
        "pageSize": "30",
        "group": "",
        "filter": "",
        "qAreaNo": "",
        "qDateBegin": "",
    }

    resp_json = requests.post(url,data=data).json()
    docs = []
#     return resp_json['Data'][:3]
    for section in resp_json['Data']:
        docs.append({
            'id':str( section['Id'] ),
            '年度場次':'{}年第{}場'.format( section['sAucYear'],section['sAucNo'] ),
            '區域':section['sAreaNo'],
#             '開標時間':section['sSaleDateTime'].replace('\u3000',' '),
            '投標日期':section['sSaleDateTW'],
            '開標時間':section['sSaleTime'],
            '投標時間':section['sBidTime'],
            'BidDateTime':section['sSaleDate']+' '+section['sBidTime'],
            'SaleDateTime':section['sSaleDate']+' '+section['sSaleTime'],
        })
    ###
    return docs

def getBuzBidTime(data):
    req_data={
        "sort": "",
        "page": "1",
        "pageSize": "100000",
        "group": "",
        "filter": "",
        "qTab": "tabgold",
        "Id": str(data['id']),
        "sCollect": "0",
    }

    url = 'https://www.tfasc.com.tw/Product/BuzBidTime/ReadDataDetail'
    resp_json = requests.post(url,data=req_data).json()
    #print(resp_json)
    docs = []
#     return resp_json['Data'][:3]
    for item in resp_json['Data']:
        if str(item['FileRoot']) != 'None':
            doc = data.copy()

            doc['bulletin_url'] = item['FileRoot']+item['FilePath']
            doc['document'] = item['FileRoot']+item['FilePath']

            doc['estate_url'] = 'https://www.tfasc.com.tw/Product/BuzRealEstate/Detail?id='+str( item['Id'] )

            doc['金服案號'] = item['sCrmNo']
            doc['number'] = item['sCrmNo']

            doc['土地地號-建號'] = item['BuzTitleType'].split('：')[-1]
            doc['address'] = item['BuzTitleType'].split('：')[-1]
            doc['parcel'] = item['BuzTitleType'].split('：')[-1]

            doc['總底價'] = item['chNewTotalMinPrice']
            doc['reserve'] = item['chNewTotalMinPrice']

            doc['SaleNo(拍次)'] = item['SaleNo']
            doc['date'] = item['SaleNo']

            doc['sHouseType'] = item['sHouseType']
            doc['remark'] = item['sHouseType']

            doc['sSaleDate'] = item['sSaleDate']
            doc['session'] = item['sSaleDate']

            item['sBatchNo'] = num_transformer(item['sBatchNo'])
            doc['標別'] = item['sBatchNo']

            doc['拍次'] = item['sNewSaleNo']

            doc['總面積(坪)(持分)'] = item['pingRrange']
            doc['最低拍賣價格'] = item['MinPrice']

            docs.append(doc)
        else:
            pass
    return docs

def getEstate(doc,proxy=''):
    if proxy: proxies = {'http':proxy,'https':proxy}
    else: proxies = {}
    resp = requests.get(doc['estate_url'],proxies=proxies)
    dom = BeautifulSoup(resp.text,'lxml')
    if '財產所有人' in dom.select('.panel-heading span')[0].text:
        doc['owner'] = dom.select('.panel-heading span')[0].text.split('：')[-1].strip('（） │|\r')
    if dom.select('#land-detail th'):
        # 縣市   鄉鎮市區   段   小段   地號   面積(m2)   權利範圍   最低拍賣價格
        for k,v in zip(dom.select('#land-detail th'),dom.select('#land-detail td')):
            doc[k.text.strip()] = v.text.strip()
    else:
        # 建物門牌   主要建材與房屋層數   面積(m2)   權利範圍   最低拍賣價格   備考
        for row in dom.select('.panel-body .detail p'):
            k,v = row.text.split('：',1)
            doc[k] = v
    return doc

###
def etl(docs):
    document=''
    number=''
    price=''
    sale_no=''
    for doc in docs:
        if doc.get('金服案號',''):
            document = doc['bulletin_url']
            number = doc['金服案號']
        if doc.get('拍次',''):
            price = doc['總底價']
            sale_no = doc['拍次']
        if doc['bulletin_url']==document:
            doc['金服案號'] = number
            doc['number'] = number
            doc['總底價'] = price
            doc['reserve'] = price
            doc['拍次'] = sale_no
    return docs

def address_split(doc):
    address = doc.get('address','')
    _parcel = []
    for key_ls in [ ['縣','市'],['鄉', '鎮', '市', '區'] ]:
        key_re = '|'.join(key_ls)
        key = ''.join(key_ls)
        sp = re.search(key_re,address)
        if sp:
            sp = sp.group()
            _sp = address.split(sp,1)
            doc[key] = _sp[0]+sp
            _parcel.append( _sp[0]+sp )
            address = _sp[1]

    sp_int = re.search('\d',address)
    if sp_int:
        sp_int = sp_int.group()
        _sp = address.split(sp_int,1)
        _sp[1] = sp_int + _sp[1]
        _parcel+=_sp
    doc['parcel'] = ' '.join(_parcel)
    return doc

def bulletin_etl(raw_tb,regexp):
    def _etl(tb,_type):
        _type = '鄉鎮市區' if _type=='land' else '屋層數'
        tmp_tb=[]
        result_tb=[]
        _start=False
        for row in tb:
            if '點交情形' in row: _start=False
            if _start:
                tmp_tb.append(row)

            if not _start and _type in row: _start=True
            if '│備考' in row:
                result_tb.append(tmp_tb)
                tmp_tb=[]
        return result_tb

    build_tb = []
    land_tb=[]
    for tb1 in raw_tb:
        _tb=re.split(regexp,tb1.split('\r\n│使用情形')[0])[0].split('\r\n')
        tmp_tb=[]
        _start=False
        _is_build=False
        for _ in _tb:
            if '點交情形' in _:
                _start=False
            if _start:
                tmp_tb.append(_)
            if not _start and '最低拍賣價格' in _:
                _start=True
            if '│備考' in _:
                if '屋層數' in ''.join(tmp_tb):
                    _is_build=True
                if _is_build:
                    build_tb+=tmp_tb
                else:
                    land_tb+=tmp_tb
                tmp_tb=[]
    build_tb = _etl(build_tb,'build')
    land_tb = _etl(land_tb,'land')
    return [land_tb,build_tb]

def split_owner(raw_string):
    owners_temp = raw_string.split('財產所有人：')[-1].replace('兼',',').replace('。',',').replace('/n',',').replace('、',',').replace('，',',').replace(' ',',').replace('：',',').replace('歿',',').replace('與',',').replace('即',',').replace('之繼承人',',').replace('繼承人',',').replace('之遺產管理人',',').replace('遺產管理人',',').replace(':',',').replace('之限定繼承人',',').replace('限定繼承人',',').replace('原名',',').replace('權利範圍',',').replace('均為',',').replace('應有',',').replace('之有限責任繼承人',',').replace('有限責任繼承人',',').replace('之律師',',').replace('律師',',').replace('之清算管理人',',').replace('清算管理人',',').replace("（",',').replace("）",',').replace("(",',').replace(")",',').replace('之遺產管理人',',').replace('遺產管理人',',').replace('〈',',').replace('〉',',').replace('之限定',',').split(',')
    owners_temp = list(set(filter(None, owners_temp)))
    if len(owners_temp) > 1:
        owners_org = raw_string.split('財產所有人：')[-1]
    else:
        owners_org = ''
    return owners_temp,owners_org
    
def getOwner(raw_tb,regexp):
    res_dict = {}
    for tb in raw_tb:
        if re.search(r'^.\r\n', tb):
            _num = '標別：'+re.search(r'^.\r\n', tb).group().strip()
        else:
            _num=''
        tmp_tb = []
        _start=False
        _tb = re.split(regexp,tb.split('\r\n│使用情形')[0])[0].split('\r\n')
        for row in _tb:
            if re.search(r'^├',row):
                _start=False
                break
            if '財產所有人：' in row:
                _start=True
            if _start:
                tmp_tb.append(row)
        _num = num_transformer(_num)
        _res = ''.join([_.strip('│\u3000 ') for _ in tmp_tb])
        res_dict[_num] = '、'.join( split_owner(_res)[0] )
        ###
        if res_dict[_num]=='如備考欄所載':
            _tmp = []
            _start=False
            for _ in tb.split('\r\n'):
                if '│備考' in _: _start=True
                if _start:
                    if re.search(r'^├',_):
                        _start=False
                        break
                    _tmp.append(_.replace(' ','').replace('、','，'))
            _res = _tmp[0].split('│')
            for _t in _tmp[1:]:
                for i in range(len(_res)):
                    _res[i] = (_res[i]+_t.split('│')[i]).strip()
            res_dict[_num] = '、'.join( split_owner(_res[3])[0] )
        ###
    return res_dict,split_owner(_res)[1]

def getBulletin(document,number):
    resp = requests.get(document)
    resp.encoding='big5'
    resp_text = resp.text
    content,doc_tb = resp_text.split('附表：')
    regexp="{}（(.+?)）".format(re.sub(r'\d+','\d+',number))
    res=''.join([_.strip() for _ in content.split('\r')])
    print(res)
    ###
    try:
        try:
            court,court_number = re.search(regexp,res).group(1).split('案號：')
        except:
            court,court_number = re.search(regexp,res).group(1).split('：')
    except:
        pass

    raw_tb = [_ for _ in re.split('標別：',doc_tb) if len(_)>100]
    owner_dict = getOwner(raw_tb,regexp)[0]
    owners_org = getOwner(raw_tb,regexp)[1]
    return bulletin_etl(raw_tb,regexp)+[court,court_number,owner_dict,owners_org]

def get_build_number(build_tb):
    unit_list = []
    for tb in build_tb:
        _start=False
        result=[]
        for row in tb:
            if '------' in row:
                _start=False
            if _start:
                result.append( row )
            if re.search(r'^├',row):
                _start=True

        _tmp = result[0].split('│')
        for res in result[1:]:
            for i in range(len(_tmp)):
                _tmp[i] = (_tmp[i]+res.split('│')[i]).strip()
        unit_list.append(_tmp[2])
    return unit_list