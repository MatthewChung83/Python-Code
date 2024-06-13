from datetime import datetime, timedelta, date

import imaplib
import email

from email.header import decode_header
# pip install imap-tools
from imap_tools.imap_utf7 import encode, decode

ACCOUNT = '10773016@gm.scu.edu.tw'
PWD = '00007740'

mail = imaplib.IMAP4_SSL('imap.gmail.com')
mail.login(ACCOUNT, PWD)
status, data = mail.list()
mail.select('inbox')
# 取前一天的時間
today = (date.today() - timedelta(1)).strftime('%d-%b-%Y')
# 找出前一天未讀的信件
typ, data = mail.search(None, '(UNSEEN)', '(SENTSINCE {0})'.format(today))

# 取到 message id list(message id 就是每封信的 id)
ids = data[0]
id_list = ids.split()

for i in id_list:
    typ, data = mail.fetch(i, '(RFC822)')
    if typ == 'OK':
        print('get mail ok')

    for response_part in data:
            if isinstance(response_part, tuple):
                rawMail = response_part[1].decode()
                # 取得信件
                msg = email.message_from_string(rawMail)
                print(msg)