from os import listdir
from os.path import isfile, isdir, join
import pandas as pd

# 指定要列出所有檔案的目錄
path = r"C:\Py_Project\output\\"
comnm = 'Estate_output_'

# 取得所有檔案與子目錄名稱
files = listdir(path)
names = locals()
text = {}
for i in range(len(files)):
    if comnm in files[i]:
        names['file%s' % i] = pd.read_csv(path+files[i]).drop_duplicates()
        names['file%s' % i] = names['file%s' % i].query("psid != 'psid'").drop('data_date',axis=1)
        
        if i == 0:
            text = names['file%s' % i]
        else:
            text = text.append(names['file%s' % i])
            
text = text.drop_duplicates()
text = text.query("psid != 'psid'")
text.reset_index(drop=True)