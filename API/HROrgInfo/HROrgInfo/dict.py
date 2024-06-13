# -*- coding: utf-8 -*-
"""
Created on Mon May  9 18:29:00 2022

@author: admin
"""


def NewCash(doc):
    NewCash_Emp= []

    NewCash_Emp.append({
        
        'SYS_VIEWID':doc[0],
        'SYS_NAME':doc[1],
        'SYS_ENGNAME':doc[2],
        'TMP_PDEPARTID':doc[3],
        'TMP_PDEPARTNAME':doc[4],
        'TMP_PDEPARTENGNAME':doc[5],
        'SYS_ID':doc[6],
        'TMP_MANAGERID':doc[7],
        'TMP_MANAGERNAME':doc[8],
        'TMP_MANAGERENGNAME':doc[9],

    })
    return NewCash_Emp