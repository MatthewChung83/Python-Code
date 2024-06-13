#Email 套件 + Decode 解碼套件
import mailbox
from email.header import decode_header
#日期套件
from datetime import datetime
#os層套件
import os
#json套件
import json
#requests & post
import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

#def 設定
subject_TTL = []
def decode_mime_words(s):
    return ' '.join(word.decode(charset or 'utf-8') if charset else word
                    for word, charset in decode_header(s))

import configparser

config = configparser.ConfigParser()
#config.read('config.ini')
with open('D:\CL_Test\config.ini', 'r', encoding='utf-8') as file:
    config.read_file(file)
# 使用配置
print(config['Mail_Config']['Mail'])
#Mail路徑讀取
Mail_path= config['Mail_Config']['Mail']


def read_mbox_subjects(file_path):
    mbox = mailbox.mbox(file_path)
    for message in mbox:
        subject = message['subject']
        decoded_subject = decode_mime_words(subject)
        #print("Decoded Subject:", decoded_subject)
        subject_TTL.append(decoded_subject)
    return subject_TTL


#Step1 確認今日資料夾是否已建立
date = datetime.today().strftime('%Y%m%d')
time = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
check_date = datetime.today().strftime('%Y/%m/%d')
folder_path = rf'D:/Auto_Return_Case/{date}'
try:
    os.makedirs(folder_path, exist_ok=True)
except Exception as e:
    pass 

try:
    path = folder_path + rf'/{date}.txt'
    f = open(path, 'w')
    f.write(date+'\n')
    f.close()
except :
    pass

#Step2 讀取中租單項委外Return Mail
# 替换为你的 .mbox 文件路径
mbox_file_path = Mail_path
return_case = read_mbox_subjects(mbox_file_path)
#print(return_case)
file_path = folder_path + rf'/{date}.txt'
for i in range(len(return_case)):
    if check_date in return_case[i] :
        data_to_check = return_case[i]
        print(data_to_check)
        #讀取文件並檢查數據內容
        if '取消外訪' in data_to_check:
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    contents = file.read()
                    if data_to_check in contents:
                        print(f"'{data_to_check}' 已存在于文件中。")
                    else:
                        url = 'https://192.168.200.68:5001/api/Case/Cl/SetCaseCancelCl/'
                        headers = {'Content-Type': 'application/json'}
                        data={
                                "DataDt": str(data_to_check.split('-')[0]),
                                "Casei": str(data_to_check.split('-')[2]),
                                "Status":str(data_to_check.split('-')[1])
                        }

                        data_json = json.dumps(data)
                        print(data_json)
                        response = requests.post(url, data = data_json,headers=headers,verify=False)
                        #print(response.text)
                        # 如果資料不存在則新增

                        with open(file_path, 'a', encoding='utf-8') as file_to_append:
                            file_to_append.write(data_to_check + '，'+response.text +'\n')
                        print(f"'{data_to_check}' 已添加到文件中。 " )#+response.text)
            except FileNotFoundError:
                # 如果文件不存在，創建文件並寫入錯誤
                with open(file_path, 'w', encoding='utf-8') as file_to_write:
                    file_to_write.write(f"處理文件时出错:"+data_to_check + '\n')
        else:
            continue
        
    else:
        pass
