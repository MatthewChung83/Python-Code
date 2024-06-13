from itemadapter import ItemAdapter
from datetime import datetime
from scrapy.exporters import CsvItemExporter
from scrapy import signals
from pydispatch import dispatcher
from crawler.config import paths
import pandas as pd
import time

import pymssql 
from crawler.config import db


def item_type(item):
    return type(item).__name__.replace('Item', '').lower()  # AAAItem => aaa


class SQLPipeline(object):
    def process_item(self, item, spider):
        item['entrydate'] = datetime.now()
        # self.toSQL_pyodbc([item], item['type'], db.get('server'), db.get('database'), db.get('username'), db.get('password'))
        self.toSQL(item, item['type'], db.get('server'), db.get('database'), db.get('username'), db.get('password'))
        return item

    def toSQL(self, docs, table, server, database, username, password):
        print(docs, table, server, database, username, password)
        conn = pymssql.connect(server=server, user=username, password=password, database=database)  
        cursor = conn.cursor()  

        docs['entrydate'] = docs['entrydate'].strftime("%Y/%m/%d %H:%M:%S")
        docs.pop('type', None)
        data_keys = ','.join(docs.keys())
        val = list(docs.values())
        val = [str(v) for v in val]
        data_symbols = "'"+"','".join(val)+"'"
        insert_cmd = """INSERT INTO {} ({})
        VALUES ({});""".format(table, data_keys, data_symbols)
        print(insert_cmd)
        cursor.execute(insert_cmd)  
        conn.commit()
        conn.close()

    def toSQL_pyodbc(self, docs, table, server, database, username, password):
        import pyodbc
        print(docs, table, server, database, username, password)
        conn_cmd = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password
        with pyodbc.connect(conn_cmd) as cnxn:
            cnxn.autocommit = False
            with cnxn.cursor() as cursor:
                data_keys = ','.join(docs[0].keys())
                data_symbols = ','.join(['?' for _ in range(len(docs[0].keys()))])
                insert_cmd = """INSERT INTO {} ({})
                VALUES ({})""".format(table, data_keys, data_symbols)
                data_values = [tuple(doc.values()) for doc in docs]
                cursor.executemany(insert_cmd, data_values)
                cnxn.commit()


class CrawlerPipeline(object):
    save_types = ['wbtcourtauctiontb', 'auctioninfotb']
    save_types_name = {
        'wbtcourtauctiontb': 'wbt_court_auction_tb',
        'auctioninfotb': 'auction_info_tb',
    }

    def __init__(self):
        dispatcher.connect(self.spider_opened, signal=signals.spider_opened)
        dispatcher.connect(self.spider_closed, signal=signals.spider_closed)

    def spider_opened(self, spider):
        # self.files = dict([(name, open(paths.get('output_dir')+self.save_types_name[name]+'.csv', 'ab+')) for name in self.save_types])
        self.files = {
            'wbtcourtauctiontb': open(paths.get('output_dir')+'wbt_court_auction_tb_tmp.csv', 'wb+'),
            'auctioninfotb': open(paths.get('output_dir')+f'auction_info_tb_tmp.csv', 'wb+'),
        }
        self.exporters = dict([(name, CsvItemExporter(self.files[name])) for name in self.save_types])
        [e.start_exporting() for e in self.exporters.values()]

    def spider_closed(self, spider):
        [e.finish_exporting() for e in self.exporters.values()]
        [f.close() for f in self.files.values()]
        try:
            df1 = pd.read_csv(paths.get('output_dir')+'wbt_court_auction_tb_tmp.csv')
            df2 = pd.read_csv(paths.get('output_dir')+'wbt_court_auction_tb.csv')
            result = pd.concat([df1, df1]).drop_duplicates(subset='rowid')
            result.to_csv(paths.get('output_dir')+'wbt_court_auction_tb.csv', index=False)
        except:
            pass

        try:
            df3 = pd.read_csv(paths.get('output_dir')+'auction_info_tb_tmp.csv')
            df4 = pd.read_csv(paths.get('output_dir')+'auction_info_tb.csv')
            result = pd.concat([df3, df4]).drop_duplicates(subset='rowid')
            result.to_csv(paths.get('output_dir')+'auction_info_tb.csv', index=False)
        except:
            pass
        #time.sleep(5)

    def process_item(self, item, spider):
        what = item_type(item)
        item['entrydate'] = datetime.now()
        if what in set(self.save_types):
            self.exporters[what].export_item(item)
        return item
