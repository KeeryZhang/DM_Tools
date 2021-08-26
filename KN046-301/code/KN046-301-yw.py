#!/usr/bin/python
# -*- coding:UTF-8 -*-

import openpyxl
from openpyxl.descriptors import base
from openpyxl.utils import get_column_letter, column_index_from_string
from datetime import datetime
import os
import sys
import argparse
from copy import deepcopy

sys.path.append("..\..")

from tool_lib.utils import mark, findkeyscolumn, exist, message

keys1list = ['筛选号', '访视名称', '当前使用药物']
keys2list = [r'{change}', '[Subject]', '[InstanceName]', '[EXIDEN]']


def data1(ws1, keys1):
    data_ws = {}
    for row in range(7, ws1.max_row+1):
        subject = ws1[keys1['筛选号']+str(row)].value
        instance_raw = ws1[keys1['访视名称']+str(row)].value
        drug = ws1[keys1['当前使用药物']+str(row)].value

        instancelist = instance_raw.split('D')
        instance = instancelist[0]

        drugset = set()
        if drug != None:
            druglist = drug.split(';')
            for dr in druglist:
                if dr == '':
                    continue
                name, code = dr.split(':')
                if name == '无编号药物':
                    continue
                drugset.add(code.split('S')[1])
        data_ws.setdefault(subject, {})
        data_ws[subject].setdefault(instance, {'row': row, 'drug': drugset})
    return data_ws


def data2(ws2, keys2):
    data_ws = {}
    for row in range(2, ws2.max_row+1):
        if ws2[keys2[r'{change}']+str(row)].value == 'deleted':
            continue
        if ws2[keys2['[Subject]']+str(row)].value == None:
            continue

        subject = ws2[keys2['[Subject]']+str(row)].value
        instance = ws2[keys2['[InstanceName]']+str(row)].value
        EXIDEN = ws2[keys2['[EXIDEN]']+str(row)].value

        EXIDENset = set()
        if EXIDEN != None:
            EXIDENlist = EXIDEN.split(';')
            for line in EXIDENlist:
                EXIDENset.add(line)
        
        data_ws.setdefault(subject, {})
        data_ws[subject].setdefault(instance, {'row': row, 'drug': EXIDENset})
    return data_ws


def visitcheck(data_ws1, data_ws2, ws1):
    ws1.insert_cols(1)
    ws1['A6'].value = 'visit->研究药物给药核查结果'

    for id in data_ws1:
        pid_ws1 = data_ws1[id]
        if id not in data_ws2:
            iderror = True
        else:
            pid_ws2 = data_ws2[id]
            iderror = False

        for instance in pid_ws1:
            ipid_ws1 = pid_ws1[instance]
            msg = ''
            if 'C' not in instance:
                visitpass = True
            else:
                visitpass = False

            if not visitpass:    
                if not iderror:
                    if instance not in pid_ws2:
                        instanceerror = True
                    else:
                        ipid_ws2 = pid_ws2[instance]
                        instanceerror = False

            if visitpass:
                rsg = 'Info: 该行访视为 {} 无需匹配'.format(instance)
            elif iderror:
                rsg = 'Error: 该受试者{}信息在研究药物给药页面不存在'.format(id)
            elif instanceerror:
                rsg = 'Error: 该受试者{}的访视{}信息在研究药物给药页面不存在'.format(id, instance)
            else:
                if ipid_ws1['drug'] == ipid_ws2['drug']:
                    rsg = 'Info: 该行使用药物匹配成功'
                else:
                    ws1diffws2 = ipid_ws1['drug'].difference(ipid_ws2['drug'])
                    ws2diffws1 = ipid_ws2['drug'].difference(ipid_ws1['drug'])
                    rsg = 'Error: 该行使用药物匹配失败，受试者 {0} 访视 {1}'.format(id, instance)            
                    if len(ws1diffws2) != 0:
                        rsg += ' visit页面多出 {}'.format(' '.join(str(x) for x in ws1diffws2))
                    if len(ws2diffws1) != 0:
                        rsg += ' 给药页面多出 {}'.format(' '.join(str(x) for x in ws2diffws1))

            msg = message(msg, rsg)
            mark(ws1, 'A', ipid_ws1['row'], msg)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--visit", default=r'Copy of Subject_visit_report_KN046-3011629948672870.xlsx', help="Please add AE file name")
    parser.add_argument("--gy", default=r'KN046-301_研究药物给药.xlsx', help="Please set sheet name of ae")
    parser.add_argument("--ssz", default=r'受试者', help="Please set sheet name of cb")
    parser.add_argument("--ywgy", default=r'EX|研究药物给药', help="Please set sheet name of cb")

    args = parser.parse_args()

    visit_path = os.path.join(r'..\sheets', args.visit)
    gy_path = os.path.join(r'..\sheets', args.gy)

    ssz_sheet = args.ssz
    ywgy_sheet = args.ywgy

    visit_pathlist = visit_path.split('.xlsx')

    wbsavepath = ''.join([''.join([visit_pathlist[0], '_checkout']), '.xlsx'])
    try:
        wb1 = openpyxl.load_workbook(visit_path)
        ws1 = wb1[ssz_sheet]
        
        wb2 = openpyxl.load_workbook(gy_path)
        ws2 = wb2[ywgy_sheet]
               
        keys1 = {'筛选号': 'C', '访视名称': 'L', '当前使用药物': 'R'}
        keys2 = findkeyscolumn(ws2, keys2list)

        data_ws1 = data1(ws1, keys1)
        data_ws2 = data2(ws2, keys2)
        
        visitcheck(data_ws1, data_ws2, ws1)
        # gycheck(data_ws1, data_ws2, ws2)

        wb1.save(wbsavepath)

    finally:
        wb1.close()
        wb2.close()