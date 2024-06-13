# -*- coding: utf-8 -*-
"""
Created on Wed Nov  2 14:31:03 2022

@author: admin
"""

import os
import shutil
try:
 
    file_source = rf'C:\Py_Project\project\judicial_cdcb3\save_file\\'
    file_destination = rf'\\fortune\Cashfile\UCS\judicial_cdcb3\\'
     
    get_files = os.listdir(file_source)
     
    for g in get_files:
        shutil.copy(file_source + g, file_destination)
   
    
except:
    #ERR_mail(obs)
    pass