python version: 3.6, 3.7
Usage: python request.py

dict.py --- 引用之詞庫檔:
  查詢關鍵字範圍(qtypes) 
  斷章詞庫(segwords) 
  擷句詞庫(splits) 
  切詞詞庫(delimis/delws)

etl_func.py --- 引用之功能檔 :
  src_obs: 查詢資料庫是否已存在當前訪問之法文(判斷欄位：judicial_no | judicial_doc_no | publishing_date)
  Parse_Title: 標題擷句
  Parse_Subject: 主旨擷句
  Parse_anncm: 公告擷句
  delws_re: 擷句後切詞
  indata: 法文主檔入庫前格式彙整
  indetail: 債務人檔入庫前格式彙整
  toSQL: 寫入資料庫

爬蟲主檔 request.py 步驟說明：
  1. call API
  2. 確認查詢有無結果
  3. 分頁爬取
  4. 分案爬取
  5. 解析(斷章 + 擷句 + 切詞)
  6. 確認當前法文不存在資料庫
  7. 法文主檔與債務人檔入庫

