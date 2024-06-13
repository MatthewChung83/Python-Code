# -*- coding: utf-8 -*-
"""
Created on Sat Sep 11 14:14:27 2021

@author: MATTHEW5043
"""

"第一種作法"
#%%"進行虛擬空間安裝"--cmd --winsowspowershell
pip install virtualenv 
#%%使用--cmd or --winsowspowershell 創建venv空間
virtualenv venv
#%%將需要用到模組輸入至requirements.txt
pip freeze > requirements.txt
#%%下載requirements.txt
pip install -r requirements.txt
#%%離開虛擬環境空間
deactivate
#%%結束

"第二種作法 cmd"
#%%安裝套件
easy_install virtualenv
#%%建立虛擬環境
virtualenv [指定虛擬環境的名稱]
#%%啟動虛擬環境
D:\spyder\ENV\Scripts\activate
#%%
cd /d 至ENV\Lib
#%%

