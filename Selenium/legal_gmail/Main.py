import imaplib
import email
from email.header import decode_header
import re

# 登錄信息
username = 'ucsin001'
password = '23756020'

# 建立與 Gmail 的連接
mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login(username, password)

# 選擇郵箱，'INBOX' 表示收件箱
mail.select('inbox')

# 搜索郵件
status, messages = mail.search(None, '(SUBJECT "驗證碼")')
# 如果要讀取所有郵件，可以將條件改為 'ALL'

if status != 'OK':
    print("沒有找到郵件！")
else:
    # messages 是一個郵件編號列表，這裡只取最新的一封
    latest_email_id = messages[0].split()[-1]
    status, data = mail.fetch(latest_email_id, '(RFC822)')

    if status != 'OK':
        print("無法讀取郵件！")
    else:
        # 解析郵件內容
        msg = email.message_from_bytes(data[0][1])
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain" or part.get_content_type() == "text/html":
                    message = part.get_payload(decode=True).decode()
                    # 使用正則表達式尋找驗證碼
                    match = re.search(r'驗證碼為 (\d{6})', message)
                    if match:
                        print('找到的驗證碼:', match.group(1))
                        break
        else:
            # 非多部分郵件，直接搜尋驗證碼
            message = msg.get_payload(decode=True).decode()
            match = re.search(r'驗證碼為 (\d{6})', message)
            if match:
                print('找到的驗證碼:', match.group(1))

# 斷開連接
mail.logout()
