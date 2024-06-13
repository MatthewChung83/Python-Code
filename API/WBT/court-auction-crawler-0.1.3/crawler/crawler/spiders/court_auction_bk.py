import json
import logging
import re
import tempfile
from datetime import datetime, timedelta
import scrapy
from crawler.items import AuctionInfoTbItem, WbtCourtAuctionTbItem
from crawler.utils import generate_id, parse_tw_date
from dateutil import parser
from pdftool.model import PDFReader
import glob
from crawler.config import paths, db
#from crawler.exist_db_file import exist_db

# add initial & incremental src, add by Chris 2021/3/1
def exist_db():
    import pymssql
    conn = pymssql.connect(server='vnt07', user='pyuser', password='Ucredit7607', database='UIS')
    cursor = conn.cursor()
    script = f"""select distinct CONCAT(court,'_',number,'_',REPLACE(REPLACE(date,' ','_'),'/',''),'.pdf') from wbt_court_auction_tb"""
    cursor.execute(script)
    qry = cursor.fetchall()
    cursor.close()
    conn.close()
          
    exist_pdf = []        
    for i in range(len(qry)):
        pdf = qry[i][0]
        exist_pdf.append(pdf)
    return exist_pdf

logging.getLogger("pdfminer").setLevel(logging.WARNING)


class CourtAuction(scrapy.Spider):
    name = 'court_auction'
    custom_settings = {
        'COOKIES_ENABLED': False,
        'DOWNLOAD_DELAY': 0.1,
    }

    def __init__(self, start_date=f"{datetime.now().year}{datetime.now().strftime('%m')}{datetime.now().strftime('%d')}", **kwargs):
        """ init
        Args:
            start_date (string): 爬蟲開始日期
        """
        self.start_date = start_date
        super().__init__(**kwargs)
        self.existed_file = glob.glob(paths.get('output_dir')+'*.pdf')
        self.existed_file = [f.split('/')[-1] for f in self.existed_file]

    def start_requests(self):
        """ 各分類爬取
        Returns:
            response (scrapy.response): 各頁爬取
        """
        date = parser.parse(self.start_date)
        start_date_str = f"{date.year-1911}{date.strftime('%m')}{date.strftime('%d')}"
        end_date = date + timedelta(days=90	)
        end_date_str = f"{end_date.year-1911}{end_date.strftime('%m')}{end_date.strftime('%d')}"

        # 一般程序  應買公告  拍定價格
        saletypes = ['1', '4', '5']
        # 房屋  土地  房屋+土地
        proptypes = ['C52', 'C51', 'C103']
        for saletype in saletypes:
            for proptype in proptypes:
                formdata = {
                    'crtnm': '全部',
                    'proptype': proptype,
                    'saletype': saletype,
                    'sorted_column': 'A.CRMYY, A.CRMID, A.CRMNO, A.SALENO, A.ROWID',
                    'sorted_type': 'ASC',
                    'pageNum': '1',
                    'saledate1': start_date_str,
                    'saledate2': end_date_str,
                    'pageSize': '100',
                }
                yield scrapy.FormRequest(
                    'https://aomp109.judicial.gov.tw/judbp/wkw/WHD1A02/QUERY.htm',
                    formdata=formdata,
                    meta={'saletype': saletype, 'proptype': proptype, 'saledate1': start_date_str, 'saledate2': end_date_str},
                    callback=self.parse_page,
                )

    def parse_page(self, response):
        """ 各頁爬取
        Args:
            response (scrapy.response)
        Returns:
            response (scrapy.response): 爬取網頁 table 內容
        """
        data_dict = json.loads(response.text)
        total_num = data_dict.get('pageInfo').get('totalNum')
        max_page = int(total_num/100) + 1
        saletype = response.meta.get('saletype')
        proptype = response.meta.get('proptype')

        for i in range(1, max_page+1):
            formdata = {
                'crtnm': '全部',
                'proptype': proptype,
                'saletype': saletype,
                'sorted_column': 'A.CRMYY, A.CRMID, A.CRMNO, A.SALENO, A.ROWID',
                'sorted_type': 'ASC',
                'pageNum': str(i),
                'saledate1': response.meta.get('saledate1'),
                'saledate2': response.meta.get('saledate2'),
                'pageSize': '100',
            }
            yield scrapy.FormRequest(
                'https://aomp109.judicial.gov.tw/judbp/wkw/WHD1A02/QUERY.htm',
                formdata=formdata,
                callback=self.parse_data,
                meta={'saletype': saletype, 'proptype': proptype},
                dont_filter=True,
            )

    def parse_data(self, response):
        """ 爬取網頁 table 內容
        Args:
            response (scrapy.response)
        Returns:
            WbtCourtAuctionTbItem (scrapy.items): 儲存 WbtCourtAuctionTbItem 至 csv
            response (scrapy.response): 擷取 pdf 內容
        """
        data_dict = json.loads(response.text)
        data = data_dict.get('data')

        saletype = response.meta.get('saletype')
        proptype = response.meta.get('proptype')

        for d in data:
            pdf_url = f'https://aomp109.judicial.gov.tw/judbp/wkw/WHD1A02/DO_VIEWPDF.htm?filenm={d.get("filenm")}'

            fields_to_pass = {
                'country': d.get('hsimun'),
                'town': d.get('ctmd'),
                'address': d.get('budadd'),
            }
            fields = {
                'court': d.get('crtnm'),
                'number': d.get('crm') + d.get('dptstr'),
                'unit': d.get('dptstr')[1:-1],
                'date': f"{d.get('saledate')[:-4]}/{d.get('saledate')[-4:-2]}/{d.get('saledate')[-2:]} 第{d.get('saleno')}拍",
                'turn': f"第{d.get('saleno')}拍",
                'reserve': d.get('saleamtstr'),
                'price': '',
            }
            key = {
                'number': d.get('crm') + d.get('dptstr'),
                'date': f"{d.get('saledate')[:-4]}/{d.get('saledate')[-4:-2]}/{d.get('saledate')[-2:]} 第{d.get('saleno')}拍",
            }
            output = {
                **key,
                **fields,
                'rowid': generate_id(key),
                'address': ' '.join([d.get('hsimun'), d.get('ctmd'), d.get('budadd'), d.get('secstr', ''), d.get('area3str'), d.get('saleamtstr')]).strip(),
                'date_datetime': parse_tw_date(f"{d.get('saledate')[:-4]}/{d.get('saledate')[-4:-2]}/{d.get('saledate')[-2:]}"),
                'country': ' '.join([d.get('hsimun'), d.get('ctmd')]),
                'reserve_int': int(d.get('minprice')),
                'handover': d.get('checkynstr'),
                'vacancy': d.get('emptyynstr'),
                'target': d.get('batchno'),
                'document': pdf_url,
                'saletype': saletype,
                'proptype': proptype,
                'remark': d.get('rmk'),
                'type': db.get('wbt_court_auction_tb'),
            }
            
            # initial and incremental
            
            pdf_file_name = '_'.join([output.get('court'), output.get('number'), d.get('saledate'), f"第{d.get('saleno')}拍"])+'.pdf'            
            existed_pdf = exist_db()
            
            if pdf_file_name not in existed_pdf:
                yield WbtCourtAuctionTbItem(output)

            # pdf_file_name = '_'.join([output.get('court'), output.get('number'), d.get('saledate'), f"第{d.get('saleno')}拍"])+'.pdf'            
            # existed_pdf = exist_db()

            #if pdf_file_name not in self.existed_file:
            if pdf_file_name not in existed_pdf:
                yield scrapy.Request(
                    pdf_url,
                    meta={
                        'fields_to_pass': fields_to_pass,
                        'fields': fields,
                        'rowid': output.get('rowid'),
                        'filename': '_'.join([output.get('court'), output.get('number'), d.get('saledate'), f"第{d.get('saleno')}拍"])+'.pdf'
                    },
                    dont_filter=True,
                    callback=self.parse_pdf,
                )

    def parse_pdf(self, response):
        """ 擷取 pdf 內容
        Args:
            response (scrapy.response)
        Returns:
            AuctionInfoTbItem (scrapy.items): 儲存 AuctionInfoTbItem 至 csv
        """
        file_path = paths.get('output_dir') + response.meta.get('filename')
        with open(file_path, 'wb') as f:
            f.write(response.body)

        report = PDFReader(file_path)
        data = report.extract()
        for d in data:
            output = {
                **d,
                **response.meta.get('fields'),
                **response.meta.get('fields_to_pass'),
                'refertb': db.get('wbt_court_auction_tb'),
                'type': db.get('auction_info_tb'),
                'referi': response.meta.get('rowid')
            }
            output = {
                'rowid': generate_id(output),
                **output,
            }
            yield AuctionInfoTbItem(output)
