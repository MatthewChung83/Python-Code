#!/usr/bin/env python
# -*- coding: utf-8 -*-
import re
from datetime import datetime

def auction_info_owner_tb_etl(docs):
    ### auction_info_owner_tb
    auction_info_owner_tb_result = []
    index=1
    
    for doc in docs:
        owner_ls = [_.strip('（） │|\r') for _ in doc['owner'].split('、') if _.strip('（） │|\r')]
        for owner in owner_ls:
            auction_info_owner_tb_result.append({
                "rowid":index,
                "auction_info_i":int(doc['estate_url'].split('id=')[1]),
                "owner": re.sub(r'（|）|\(|\)','', owner.strip('()（） │|\r')),
                "owners_org": doc['owners_org'],
            })
            index+=1
    return auction_info_owner_tb_result

def auction_info_tb_etl(docs):
    ### auction_info_tb
    auction_info_tb_result = []
    index=1
    for doc in docs:
        owner_ls = [_.strip('（） │|\r') for _ in doc['owner'].split('、') if _.strip('（） │|\r')]
        for owner in owner_ls:
            auction_info_tb_result.append({
                "rowid": index,
                "court": doc['court'],
                "number": doc['number'], #"number": doc['court_number'],
                "unit": doc.get('unit',''),
                "date": doc['投標日期'],
                "turn": doc['拍次'],
                "country": doc.get('縣市',''),
                "town": doc.get('鄉鎮市區',''),
                "address": doc['address'],
                "area": doc['總面積(坪)(持分)'],
                "reserve": doc['reserve'],
                "price": doc.get('最低拍賣價格',''),
                "remark": doc['remark'],
                "owners_org": doc['owners_org'],
                "owner": re.sub(r'（|）|\(|\)','', owner.strip('()（） │|\r')),
                "entrydate":str(datetime.now())[:-3],
                "refertb": "wbt_tfasc_auction_tb",
                "referi": int(doc['estate_url'].split('id=')[1]),
                "parcel": doc['parcel'], #
            })
            index+=1
    return auction_info_tb_result

def wbt_tfasc_auction_tb_etl(docs):
    # wbt_tfasc_auction_tb
    wbt_tfasc_auction_tb_result = []
    index=1
    for doc in docs:
        wbt_tfasc_auction_tb_result.append({
#             "rowid":index,
            "rowid":int(doc['estate_url'].split('id=')[1]),
            "session":doc['session'],
            "date":doc['拍次'],
            "number":doc['number'],
            "address":doc['address'],
            "reserve":doc['reserve'],
            "deposit":"NULL",
            "price":"NULL",
            "target":"NULL",
            "auction":"NULL",
            "remark":doc['remark'],
            "document":doc['document'],
            "entrydate":str(datetime.now())[:-3],
        })
        index+=1
    return wbt_tfasc_auction_tb_result
###
