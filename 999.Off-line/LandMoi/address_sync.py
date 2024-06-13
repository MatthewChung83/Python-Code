# csv Package
import csv
# Pandas for data clean
import pandas as pd

c_src = pd.read_excel(r'C:\Users\cyche\Documents\Python Scripts\land_moi\address_identification.xlsx',sheetname='轉換後地址清單').drop_duplicates()
c_index = pd.read_csv(r'C:\Users\cyche\Documents\Python Scripts\land_moi\address_index.csv').drop_duplicates()

file_path = r'C:\Users\cyche\Documents\Python Scripts\land_moi\testing_ds_output.csv'

def write_to_csv(path,datalist):
    csvFileToWrite = open(path, 'a', encoding='utf-8-sig', newline='')
    csvDataToWrite = csv.writer(csvFileToWrite)
    csvDataToWrite.writerow(datalist)
    csvFileToWrite.close()

def addr_check(i):
    src_casei = c_src['casei'][i]
    src_address = c_src['address'][i]
    
    src_addr_class = c_src['addr_class'][i]
    src_confirm_status = c_src['confirm_status'][i]
    src_ID = c_src['ID'][i]
    
    src_city = c_src['city'][i]
    src_town = c_src['town'][i]
    src_road = c_src['road'][i]
    src_lane = c_src['lane'][i]
    src_alley = c_src['alley'][i]
    src_number = c_src['number'][i]
    src_miss = c_src['miss_info'][i]
    src_data = None
    index_data1 = None
    index_data2 = None
    result = None
    
    src_data = [[src_casei],[src_address],[src_city],[src_town],[src_road],[src_lane]]
    index_data1 = c_index.loc[c_index["city"] == src_city].loc[c_index["town"] == src_town].loc[c_index["road"] == src_road]
    index_data2 = c_index.loc[c_index["city"] == src_city].loc[c_index["town"] == src_town].loc[c_index["road"] == src_lane]
    index_data3 = c_index.loc[c_index["city"] == src_city].loc[c_index["town"] == src_town].loc[c_index["road"] == src_miss]
    result = len(index_data1) + len(index_data2) + len(index_data3)
    
    if len(index_data2) == 1:
        src_data = [i,src_casei,src_address,src_addr_class,src_confirm_status,src_ID,src_city,src_town,src_lane,'',src_alley,src_number,len(index_data1),len(index_data2),len(index_data3),result]
        write_to_csv(file_path,src_data)
    elif len(index_data3) == 1:
        src_data = [i,src_casei,src_address,src_addr_class,src_confirm_status,src_ID,src_city,src_town,src_miss,'',src_alley,src_number,len(index_data1),len(index_data2),len(index_data3),result]
        write_to_csv(file_path,src_data)        
    else:
        src_data = [i,src_casei,src_address,src_addr_class,src_confirm_status,src_ID,src_city,src_town,src_road,src_lane,src_alley,src_number,len(index_data1),len(index_data2),len(index_data3),result]
        write_to_csv(file_path,src_data)        
    return src_data,len(index_data1),len(index_data2),len(index_data3),result

for i in range(len(c_src)):
    print(addr_check(i))