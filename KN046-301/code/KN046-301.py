#!/usr/bin/python
# -*- coding:UTF-8 -*-

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from datetime import datetime
import os
import sys
import argparse
from copy import deepcopy

sys.path.append("..\..")

from tool_lib.utils import mark, findkeyscolumn, exist, message, parse_dmy

keys1list = [r'{change}', '[Subject]', '[InstanceName]', '[TLYN]', '[TLDIAT]', '[TLDAT]', '[TLMETHOD]', '[TLLNKID]']
keys2list = [r'{change}', '[Subject]', '[InstanceName]', '[NTLYN]', '[NTLDAT]', '[NTLORRES]', '[NTLLNKID]', '[NTLMTHOD]']
keys3list = [r'{change}', '[Subject]', '[InstanceName]', '[NWTLEYN]']
keys4list = [r'{change}', '[Subject]', '[InstanceName]', '[RSYN]', '[RSDAT]', '[TRGRESP]', '[NTRGRESP]', '[NEWLIND]']


# def data1(ws, keys):
#     data_ws = {}
#     for row in range(2, ws.max_row+1):
#         if ws[keys[r'{change}']+str(row)].value == 'deleted':
#             continue
#         if ws[keys['[Subject]']+str(row)].value == None:
#             continue

#         Subject = ws[keys['[Subject]']+str(row)].value
#         InstanceName = ws[keys['[InstanceName]']+str(row)].value
#         TLYN = ws[keys['[TLYN]']+str(row)].value
#         TLDIAT = ws[keys['[TLDIAT]']+str(row)].value
#         TLDAT = ws[keys['[TLDAT]']+str(row)].value
#         TLMETHOD = ws[keys['[TLMETHOD]']+str(row)].value
#         TLLNKID = ws[keys['[TLLNKID]']+str(row)].value
        
#         data_ws.setdefault(Subject, {})
#         data_ws[Subject].setdefault(InstanceName, {})
#         data_ws[Subject][InstanceName].setdefault(row, {'TLYN':TLYN, 'TLDIAT':TLDIAT, 'TLDAT':TLDAT, 'TLMETHOD':TLMETHOD, 'TLLNKID':TLLNKID})
#     return data_ws


# def data2(ws, keys):
#     data_ws = {}
#     for row in range(2, ws.max_row+1):
#         if ws[keys[r'{change}']+str(row)].value == 'deleted':
#             continue
#         if ws[keys['[Subject]']+str(row)].value == None:
#             continue

#         Subject = ws[keys['[Subject]']+str(row)].value
#         InstanceName = ws[keys['[InstanceName]']+str(row)].value
#         NTLYN = ws[keys['[NTLYN]']+str(row)].value
#         NTLDAT = ws[keys['[NTLDAT]']+str(row)].value
#         NTLORRES = ws[keys['[NTLORRES]']+str(row)].value
#         NTLLNKID = ws[keys['[NTLLNKID]']+str(row)].value
#         NTLMTHOD = ws[keys['[NTLMTHOD]']+str(row)].value
        
#         data_ws.setdefault(Subject, {})
#         data_ws[Subject].setdefault(InstanceName, {})
#         data_ws[Subject][InstanceName].setdefault(row, {'NTLYN':NTLYN, 'NTLDAT':NTLDAT, 'NTLORRES':NTLORRES, 'NTLLNKID':NTLLNKID, 'NTLMTHOD':NTLMTHOD})
#     return data_ws


# def data3(ws, keys):
#     data_ws = {}
#     for row in range(2, ws.max_row+1):
#         if ws[keys[r'{change}']+str(row)].value == 'deleted':
#             continue
#         if ws[keys['[Subject]']+str(row)].value == None:
#             continue

#         Subject = ws[keys['[Subject]']+str(row)].value
#         InstanceName = ws[keys['[InstanceName]']+str(row)].value
#         NWTLEYN = ws[keys['[NWTLEYN]']+str(row)].value
        
#         data_ws.setdefault(Subject, {})
#         data_ws[Subject].setdefault(InstanceName, {})
#         data_ws[Subject][InstanceName].setdefault(row, {'NWTLEYN':NWTLEYN})
#     return data_ws


# def data4(ws, keys):
#     data_ws = {}
#     for row in range(2, ws.max_row+1):
#         if ws[keys[r'{change}']+str(row)].value == 'deleted':
#             continue
#         if ws[keys['[Subject]']+str(row)].value == None:
#             continue

#         Subject = ws[keys['[Subject]']+str(row)].value
#         InstanceName = ws[keys['[InstanceName]']+str(row)].value
#         RSYN = ws[keys['[RSYN]']+str(row)].value
#         RSDAT = ws[keys['[RSDAT]']+str(row)].value
#         TRGRESP = ws[keys['[TRGRESP]']+str(row)].value
#         NTRGRESP = ws[keys['[NTRGRESP]']+str(row)].value
#         NEWLIND = ws[keys['[NEWLIND]']+str(row)].value
        
#         data_ws.setdefault(Subject, {})
#         data_ws[Subject].setdefault(InstanceName, {})
#         data_ws[Subject][InstanceName].setdefault(row, {'RSYN':RSYN, 'RSDAT':RSDAT, 'TRGRESP':TRGRESP, 'NTRGRESP':NTRGRESP, 'NEWLIND':NEWLIND})
#     return data_ws


def data(ws, keys):
    data_ws = {}
    for row in range(2, ws.max_row+1):
        tmp_keys = deepcopy(keys)
        if ws[tmp_keys[r'{change}']+str(row)].value == 'deleted':
            continue
        if ws[tmp_keys['[Subject]']+str(row)].value == None:
            continue
                
        Subject = ws[tmp_keys['[Subject]']+str(row)].value
        InstanceName = ws[tmp_keys['[InstanceName]']+str(row)].value
        
        data_ws.setdefault(Subject, {})
        data_ws[Subject].setdefault(InstanceName, {})
        data_ws[Subject][InstanceName].setdefault(row, {})

        tmp_keys.pop(r'{change}')
        tmp_keys.pop('[Subject]')
        tmp_keys.pop('[InstanceName]')

        for key in tmp_keys:
            data_ws[Subject][InstanceName][row].update({key:ws[tmp_keys[key]+str(row)].value})
    return data_ws


def bbzcheck(data_ws1, data_ws4, ws1):
    ws1.insert_cols(1)
    ws1['A1'].value = '靶病灶检查结果'

    for id in data_ws1:
        pid = data_ws1[id]
        for instance in pid:
            ipid = pid[instance]
            row_delete = []
            for row in ipid:
                ripid = ipid[row]
                if ripid['TLDAT'] == None:
                    row_delete.append(row)
                    rsg = '该行无'
            if len(row_delete) > 0:
                for dr in row_delete:
                    ipid.pop(dr)

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

        data_ws1 = data(ws1, keys1)
        data_ws2 = data(ws2, keys2)
        data_ws3 = data(ws3, keys3)
        data_ws4 = data(ws4, keys4)
        
        # data_ws1 = data(ws1, keys1)

        bbzcheck(data_ws1, data_ws4, ws1)
        fbbzcheck(data_ws2, data_ws4, ws2)
        xbzcheck(data_ws3, data_ws4, ws3)

        methodcheck(data_ws1, ws1)
        methodcheck(data_ws2, ws2)

        wb.save(wbsavepath)

    finally:
        wb.close()