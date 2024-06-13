# 1. 安裝說明

## 1-1. 安裝開發環境
1. 安裝**Python**和**pip**

	至[Python官方網站](https://www.python.org/downloads/release/python-387/)下載，需使用3.7之後之版本

2. 安裝**Virtualenv**
	`pip install virtualenv`

## 1-2. 創建虛擬環境

#### Linux

1. 選定專案資料夾
	`cd ~`

2. 創建**Virtualenv**
	`virtualenv .env`

3. 使用**Virtualenv**切換環境
	 `source ~/.env/bin/activate`

#### Windows

1. 選定專案資料夾
	`cd D:\`

2. 創建**Virtualenv**
	`virtualenv .env`

3. 使用**Virtualenv**切換環境
	`cd D:\.env\Scripts`
	`activate`

## 1-3. 安裝第三方套件

在專案目錄下執行
`pip install -r requirements.txt`


## 1-4. 執行爬蟲
進入專案 `cd court-auction-crawler/crawler/`
1. 開始時間為當日 `scrapy runspider crawler/spiders/court_auction.py`

2. 指定開始時間  `scrapy runspider crawler/spiders/court_auction.py -a start_date=2021/01/14`


# 2. Scrapy 介紹

## 2-1. 總覽



**爬蟲**的原始碼位於**crawler/crawler**，整體架構說明如下

```
└── crawler                         <- 專案目錄
    └── pdftool                     <- pdf擷取程式
    └── data                        <- pdf預設儲存路徑
    └── crawler                     <- 爬蟲目錄
        ├── spiders                 <- Scrapy的Spiders元件
        │    └── court_auction.py   <- 爬蟲主程式
        ├── items.py                <- Scrapy的Item定義
        ├── middlewares.py          <- Scrapy的Download Middlewares元件
        ├── pipelines.py            <- Scrapy的Item Pipleline元件
        ├── settings.py             <- Scrapy的預設設定
        ├── config.py               <- Scrapy的參數設定
        └── utils.py                <- 通用的工具程式
```

**爬蟲**相關參數位置：
```
指定日期 -> court_auction.py -> def start_requests(self):
位置參數 -> config.py -> paths
pdf存放位置 -> court_auction.py -> def parse_pdf(self, response):
csv存放位置 -> middlewares.py -> def spider_opened(self, spider):
```

**爬蟲**在本專案使用的是**Scrapy**框架進行開發，包括以下幾個元件：

- [**Scrapy Engine**](https://docs.scrapy.org/en/latest/topics/architecture.html#scrapy-engine)：控制所有元件間的資料傳輸

- [**Scheduler**](https://docs.scrapy.org/en/latest/topics/architecture.html#scheduler)：排程來自**Engine**的不同頁面的爬取請求

- [**Downloader**](https://docs.scrapy.org/en/latest/topics/architecture.html#downloader )：下載網頁內容和其他靜態資源，透過**Engine**傳輸至**Spiders**

- [**Spiders**](https://docs.scrapy.org/en/latest/topics/architecture.html#spiders)：寫於spiders資料夾的**類物件（Class）**，用以定義資料爬取的邏輯

- [**Item Pipeline**](https://docs.scrapy.org/en/latest/topics/architecture.html#item-pipeline)：驗證、清理並轉換從**Spiders**爬取的資料，最後傳輸至資料庫

- [**Downloader Middlewares**](https://docs.scrapy.org/en/latest/topics/architecture.html#downloader-middlewares)：攔截**Engine**至**Downloader**之間的資料流，做統一的調整

**爬蟲**元件間的資料流邏輯如下圖所示：

![](https://docs.scrapy.org/en/latest/_images/scrapy_architecture_02.png)



詳細的**Scrapy API**和程式寫法可以參考spiders目錄下的爬蟲寫法，或是查看[**Scrapy官方文件**](https://docs.scrapy.org/en/latest/)。
