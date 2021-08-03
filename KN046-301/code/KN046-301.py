#!/usr/bin/python
# -*- coding:UTF-8 -*-

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from datetime import datetime
import os
import sys
import argparse

sys.path.append("..\..")

from tool_lib.utils import mark, findkeyscolumn, exist, message, parse_dmy

keys1list = [r'{change}', '[Subject]', '[InstanceName]', '[TLYN]', '[TLDIAT]', '[TLDAT]', '[TLMETHOD]', '[TLLNKID]']
keys2list = [r'{change}', '[Subject]', '[InstanceName]', '[NTLYN]', '[NTLDAT]', '[NTLORRES]', '[NTLLNKID]', '[NTLMTHOD]']
keys3list = [r'{change}', '[Subject]', '[InstanceName]', '[NWTLEYN]']
keys4list = [r'{change}', '[Subject]', '[InstanceName]', '[RSYN]', '[RSDAT]', '[TRGRESP]', '[NTRGRESP]', '[NEWLIND]']


def data1(ws1, keys1):
    data_ws1 = {}
    
    return data_ws1


def data2(ws2, keys2):
    data_ws2 = {}
    
    return data_ws2


def data3(ws3, keys3):
    data_ws3 = {}
    
    return data_ws3


def data4(ws4, keys4):
    data_ws4 = {}
    
    return data_ws4


def bbzcheck(data_ws1, data_ws4, ws1):

    return


def fbbzcheck(data_ws2, data_ws4, ws2):

    return


def xbzcheck(data_ws3, data_ws4, ws3):

    return


def methodcheck(data_ws, ws):

    return

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--cancer", default=r'KN046-301_肿瘤评估_20210803.xlsx', help="Please add AE file name")
    parser.add_argument("--bbz", default=r'TUTL|肿瘤评价-靶病灶（RECIST 1.1）', help="Please set sheet name of ae")
    parser.add_argument("--fbbz", default=r'TUNTL|肿瘤评价-非靶病灶（RECIST 1.1）', help="Please set sheet name of cb")
    parser.add_argument("--xbz", default=r'TUNEWTL|肿瘤评价-新病灶（RECIST 1.1）', help="Please set sheet name of cb")
    parser.add_argument("--recist", default=r'RS|总体疗效评价（RECIST 1.1）', help="Please set sheet name of cb")

    args = parser.parse_args()

    cancer_path = os.path.join(r'..\sheets', args.cancer)    
    bbz_sheet = args.bbz
    fbbz_sheet = args.fbbz
    xbz_sheet = args.xbz
    recist_sheet = args.recist

    cancer_pathlist = cancer_path.split('.xlsx')

    wbsavepath = ''.join([''.join([cancer_pathlist[0], '_checkout']), '.xlsx'])
    try:
        wb = openpyxl.load_workbook(cancer_path)
        ws1 = wb[bbz_sheet]
        ws2 = wb[fbbz_sheet]
        ws3 = wb[xbz_sheet]
        ws4 = wb[recist_sheet]
               
        keys1 = findkeyscolumn(ws1, keys1list)
        keys2 = findkeyscolumn(ws2, keys2list)
        keys3 = findkeyscolumn(ws3, keys3list)
        keys4 = findkeyscolumn(ws4, keys4list)

        data_ws1 = data1(ws1, keys1)
        data_ws2 = data2(ws2, keys2)
        data_ws3 = data3(ws3, keys3)
        data_ws4 = data4(ws4, keys4)
        
        bbzcheck(data_ws1, data_ws4, ws1)
        fbbzcheck(data_ws2, data_ws4, ws2)
        xbzcheck(data_ws3, data_ws4, ws3)

        methodcheck(data_ws1, ws1)
        methodcheck(data_ws2, ws2)

        wb.save(wbsavepath)

    finally:
        wb.close()