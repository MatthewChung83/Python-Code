# csv Package
import csv
# Pandas for data clean
import pandas as pd

import pymssql

# 資料來源引入
#c_src = pd.read_excel(r'C:\Users\cyche\Documents\Python Scripts\land_moi\address_identification.xlsx',sheetname='轉換後地址清單').drop_duplicates()
offobs = 0
fetobs = 10000

conn = pymssql.connect(server='VNT07', user='.\chris', password='Ucs28289788')
cursor = conn.cursor()
cursor.execute(f"SELECT [eaid],[pid],[addressi],[casei],[City],[Town],[Street],[Sec],[Lane],[Lane2],[No],[SubNo],[Floor],[SubFloor] FROM [UIS].[dbo].[ElAddrTb] ORDER BY [eaid] OFFSET {offobs} ROWS FETCH NEXT {fetobs} ROWS ONLY")
c_src = cursor.fetchall()
cursor.close()
conn.close()

c_index = pd.read_csv(r'C:\Py_Project\project\LandMoi\address_index.csv').drop_duplicates()
file_path = r'C:\Py_Project\output\address_sync_output.csv'

# 追加寫入CSV函式
def write_to_csv(path,datalist):
    csvFileToWrite = open(path, 'a', encoding='utf-8-sig', newline='')
    csvDataToWrite = csv.writer(csvFileToWrite)
    csvDataToWrite.writerow(datalist)
    csvFileToWrite.close()

# 來源參數導入
def addr_check(i):
    src_eaid = c_src[i][0]
    src_ID = c_src[i][1]
    src_addressi = c_src[i][2]
    src_casei = c_src[i][3]
    src_city = c_src[i][4].replace('台','臺')
    src_town = c_src[i][5]
    src_Street = c_src[i][6]+c_src[i][7].replace('1段','一段').replace('2段','二段').replace('3段','三段').replace('4段','四段').replace('5段','五段').replace('6段','六段').replace('7段','七段').replace('8段','八段').replace('9段','九段')
    src_lane = c_src[i][8]
    src_alley = c_src[i][9]
    src_No = c_src[i][10]+c_src[i][11]
    src_floor = c_src[i][12]+c_src[i][13]
    src_data = None
    index_data1 = None
    result = None

    index_data1 = c_index.loc[c_index["city"] == src_city].loc[c_index["town"] == src_town].loc[c_index["road"] == src_Street]
    result = len(index_data1)
   
    src_data = [i,src_eaid,src_ID,src_addressi,src_casei,src_city,src_town,src_Street,src_lane,src_alley,src_No,src_floor,len(index_data1),result]
    write_to_csv(file_path,src_data)        
    return src_data,len(index_data1),result

LandDataTitle = ('i','eaid','ID','addressi','casei','city','town','Street','lane','alley','No','floor','mapping_flg','result')
write_to_csv(file_path,LandDataTitle)  

for i in range(len(c_src)):
#for i in range(0,10):
    print(addr_check(i))