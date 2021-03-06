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

keys1list = [r'{change}', '[Subject]', '[InstanceName]', '[TLYN]', '[TLDIAT]', '[TLDAT]', '[TLMETHOD]', '[TLLNKID]']
keys2list = [r'{change}', '[Subject]', '[InstanceName]', '[NTLYN]', '[NTLDAT]', '[NTLORRES]', '[NTLLNKID]', '[NTLMTHOD]']
keys3list = [r'{change}', '[Subject]', '[InstanceName]', '[NWTLEYN]']
keys4list = [r'{change}', '[Subject]', '[InstanceName]', '[RSYN]', '[RSDAT]', '[TRGRESP]', '[NTRGRESP]', '[NEWLIND]']


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


def bbzpretriage(data_ws, data_ws4, ws, TU):
    
    id_delete = set()
    for id in data_ws:
        pid = data_ws[id]
        instance_delete = set()
        for instance in pid:
            ipid = pid[instance]
            row_delete = []
            for row in ipid:
                msg = ''
                rsg = ''
                ripid = ipid[row]
                if exist(id, data_ws4):
                    pid_ws4 = data_ws4[id]
                    if exist(instance, pid_ws4):
                        ipid_ws4 = pid_ws4[instance]
                        if TU == '?????????' and ripid['[TLYN]'] == '???':
                            for row_ws4 in ipid_ws4:
                                ripid_ws4 = ipid_ws4[row_ws4]
                            if ripid_ws4['[RSYN]'] == '???' or 'NA' in ripid_ws4['[TRGRESP]']:
                                rsg = 'Info:??????{}??????????????????Recist????????????RSYN?????????TRGRESP???NA'.format(TU)                                
                            else:
                                rsg = 'Error:???????????? {0} ????????? {1} ???{2}?????????????????????Recist??????RSYN????????????TRGRESP??????NA'.format(id, instance, TU)
                            row_delete.append(row)
                        elif TU == '?????????' and ripid['[TLYN]'] == None:
                            rsg = 'Error:??????{}????????????'.format(TU)                            
                            row_delete.append(row)

                        if TU == '????????????' and ripid['[NTLYN]'] == '???':
                            for row_ws4 in ipid_ws4:
                                ripid_ws4 = ipid_ws4[row_ws4]
                            if ripid_ws4['[NTRGRESP]'] == '?????????????????????':
                                rsg = 'Info:??????{}?????????????????????Recist??????????????????????????????'.format(TU)
                            else:
                                rsg = 'Error:???????????? {0} ????????? {1} ???{2}?????????????????????Recist??????NTRGRESP???????????????????????????'.format(id, instance, TU)
                            row_delete.append(row)
                        elif TU == '????????????' and ripid['[NTLYN]'] == None:
                            rsg = 'Error:??????{}????????????'.format(TU)
                            row_delete.append(row)

                        if TU == '?????????' and ripid['[NWTLEYN]'] == None:
                            rsg = 'Error:??????{}????????????'.format(TU)
                            row_delete.append(row)
                    else:
                        if '?????????' not in instance:
                            rsg = 'Error:???????????????Recist??????????????? {} ??????'.format(instance)
                            instance_delete.add(instance)
                        elif TU == '????????????' and '?????????' in instance:
                            rsg = 'Info:????????????Recist????????????'
                            instance_delete.add(instance)
                else:
                    if '?????????' not in instance:
                        rsg = 'Error:???????????????Recist???????????????'                       
                    else:
                        rsg = 'Info:??????????????????????????????Recist???????????????'
                    id_delete.add(id)
                msg = message(msg, rsg)
                mark(ws, 'A', row, msg)        

            if len(row_delete) > 0:
                for row in row_delete:
                    ipid.pop(row)

            if len(ipid) == 0:
                instance_delete.add(instance)
        
        if len(instance_delete) > 0:
            for instance in instance_delete:
                pid.pop(instance)

        if len(pid) == 0:
            id_delete.add(id)
    
    if len(id_delete) > 0:
        for id in id_delete:
            data_ws.pop(id)

    return data_ws


def pid_revert(pid):
    pid_normal = {}
    pid_cc = {}
    for instance in pid:
        rows = []
        rows_cc = []
        if 'CC' in instance or 'cc' in instance:
            pid_cc.setdefault(instance, {})
            for row in pid[instance]:
                rows_cc.append(row)
                for key in pid[instance][row]:
                    pid_cc[instance].setdefault(key, pid[instance][row][key])
            pid_cc[instance].setdefault('rows',rows_cc)
        else:
            pid_normal.setdefault(instance, {})
            for row in pid[instance]:
                rows.append(row)
                for key in pid[instance][row]:
                    pid_normal[instance].setdefault(key, pid[instance][row][key])
            pid_normal[instance].setdefault('rows',rows)

    return pid_normal, pid_cc


def bbzresult(InstanceName, crossbase, crosschecklist, crossinstancelist):
    index = crossinstancelist.index(InstanceName)
    check = crosschecklist[index]
    checklist = crosschecklist[:index+1]
    checklist.append(crossbase)
    TLDIAT_min = min(checklist)
    zerocheck = True
    if TLDIAT_min != 0:
        PRcheck = (check - crossbase)/crossbase
        PDcheck = (check - TLDIAT_min)/TLDIAT_min
    else:
        zerocheck = False
        
    if zerocheck:    
        if PDcheck >= 0.2 and abs(check - TLDIAT_min) >= 5:
            result = 'PD'
        elif abs(PRcheck) >= 0.3:
            result = 'PR'
        else:
            result = 'SD'
    else:
        result = '0'
    return result

    
def bbzpidresult(pid_sorted, pid_ws4):
    crossbase = int()
    crosschecklist = list()
    crossinstancelist = list()
    for i in range(0, len(pid_sorted)):
        instance = pid_sorted[i][0]
        p_ws1 = pid_sorted[i][1]
        TLDIAT = p_ws1['[TLDIAT]']
        if '?????????' in instance:
            crossbase = TLDIAT                
        else:
            crosschecklist.append(TLDIAT)
            crossinstancelist.append(instance) 

    for i in range(0, len(pid_sorted)):
        instance = pid_sorted[i][0]
        p_ws1 = pid_sorted[i][1]
        TLDIAT = p_ws1['[TLDIAT]']
        msg = ''
        rsg = ''
        if '?????????' in instance:
            crossbase = TLDIAT
            rsg = 'Info:?????????????????????????????????'
        else:
            result = bbzresult(instance, crossbase, crosschecklist, crossinstancelist)
            for row_ws4 in pid_ws4[instance]:
                if result in pid_ws4[instance][row_ws4]['[TRGRESP]']:
                    rsg = 'Info:?????????????????????????????????'
                elif result == '0':
                    rsg = 'Warn:??????????????????????????????????????????'
                else:
                    rsg = 'Error:??????????????????????????? {}??????Recist????????? {} ???????????????'.format(result, row_ws4)
        msg = message(msg, rsg)
        for row in p_ws1['rows']:
            mark(ws1, 'A', row, msg)


def bbzpidcheck(pid_ws1_ori, pid_ws4_ori, ws1):
    pid_ws1 = deepcopy(pid_ws1_ori)
    pid_ws4 = deepcopy(pid_ws4_ori)
    
    pid_normal, pid_cc = pid_revert(pid_ws1)

    if pid_normal != {}:
        pid_normal = sorted(pid_normal.items(), key = lambda time:time[1]['[TLDAT]'])
        bbzpidresult(pid_normal, pid_ws4)
            
    if pid_cc != {}:
        pid_cc = sorted(pid_cc.items(), key = lambda time:time[1]['[TLDAT]'])
        bbzpidresult(pid_cc, pid_ws4)   
    return 


def bbzcheck(data_ws1_ori, data_ws4, ws1):
    ws1.insert_cols(1)
    ws1['A1'].value = '?????????????????????'

    data_ws1 = deepcopy(data_ws1_ori)  

    data_ws1 = bbzpretriage(data_ws1, data_ws4, ws1, '?????????')

    for id in data_ws1:
        pid_ws1 = data_ws1[id]
        pid_ws4 = data_ws4[id]
        bbzpidcheck(pid_ws1, pid_ws4, ws1)

    return


def fbbzresult(ipid):
    NTLORRES_set = set()
    result = ''
    for row_ws2 in ipid:
        ripid_ws2 = ipid[row_ws2]
        NTLORRES_set.add(ripid_ws2['[NTLORRES]'])
    if '???????????????' in NTLORRES_set:
        result = 'PD'
    elif '??????' in NTLORRES_set:
        result = r'???CR/???PD'
    elif '????????????' in NTLORRES_set:
        result = 'NE'
    elif len(NTLORRES_set) > 0 and len(NTLORRES_set.union({'?????????'})) == 1:
        result = 'CR'
    elif len(NTLORRES_set) == 0:
        result = '0'
    return result


def fbbzpidcheck(pid_ws2_ori, pid_ws4_ori, ws2):
    pid_ws2 = deepcopy(pid_ws2_ori)
    pid_ws4 = deepcopy(pid_ws4_ori)
    
    for instance in pid_ws2:
        ipid_ws2 = pid_ws2[instance]
        ipid_ws4 = pid_ws4[instance]
        result = fbbzresult(ipid_ws2)
        msg = ''
        rsg = ''
        for row_ws4 in ipid_ws4:
            ripid_ws4 = ipid_ws4[row_ws4]
            if result in ripid_ws4['[NTRGRESP]']:
                rsg = 'Info:???????????????????????????????????????'
            elif result == '0':
                rsg = 'Error:??????????????????????????????'
            else:
                rsg = 'Error:?????????????????????????????? {} ???Recist????????? {} ????????????????????????'.format(result, ripid_ws4['[NTRGRESP]'])
        msg = message(msg, rsg)

        for row_ws2 in ipid_ws2:
            ripid_ws2 = ipid_ws2[row_ws2]
            if ripid_ws2['[NTLORRES]'] == None:
                msg = 'Error:??????????????????????????????'
            mark(ws2, 'A', row_ws2, msg)
    return


def fbbzcheck(data_ws2_ori, data_ws4, ws2):
    ws2.insert_cols(1)
    ws2['A1'].value = '????????????????????????'

    data_ws2 = deepcopy(data_ws2_ori)  

    data_ws2 = bbzpretriage(data_ws2, data_ws4, ws2, '????????????')

    for id in data_ws2:
        pid_ws2 = data_ws2[id]
        pid_ws4 = data_ws4[id]
        fbbzpidcheck(pid_ws2, pid_ws4, ws2)
    return


def xbzpidcheck(pid_ws3, pid_ws4, ws3):
    for instance in pid_ws3:
        ipid_ws3 = pid_ws3[instance]
        ipid_ws4 = pid_ws4[instance]
        for row_ws3 in ipid_ws3:
            msg = ''
            rsg = ''
            ismatch = False
            ripid_ws3 = ipid_ws3[row_ws3]
            for row_ws4 in ipid_ws4:
                ripid_ws4 = ipid_ws4[row_ws4]
            if ripid_ws3['[NWTLEYN]'] == '???':
                if ripid_ws4['[NEWLIND]'] == '???':
                    ismatch = True
            elif ripid_ws3['[NWTLEYN]'] == '???':
                if ripid_ws4['[NEWLIND]'] == '???':
                    ismatch = True
            
            if ismatch:
                rsg = 'Info:????????????????????????????????????'
            else:
                rsg = 'Error:??????????????????????????? {} ???Recist????????? {} ????????????????????????'.format(ripid_ws3['[NWTLEYN]'], ripid_ws4['[NEWLIND]'])
            msg = message(msg, rsg)
            mark(ws3, 'A', row_ws3, msg)
    return


def xbzcheck(data_ws3_ori, data_ws4, ws3):
    ws3.insert_cols(1)
    ws3['A1'].value = '?????????????????????'

    data_ws3 = deepcopy(data_ws3_ori)  

    data_ws3 = bbzpretriage(data_ws3, data_ws4, ws3, '?????????')

    for id in data_ws3:
        pid_ws3 = data_ws3[id]
        pid_ws4 = data_ws4[id]
        xbzpidcheck(pid_ws3, pid_ws4, ws3)
    return


def methodpretriage(data_ws, ws, TU):
    if TU == '?????????':
        idname = '[TLLNKID]'
        method = '[TLMETHOD]'
        ifcheck = '[TLYN]'
    else:
        idname = '[NTLLNKID]'
        method = '[NTLMTHOD]'
        ifcheck = '[NTLYN]'

    data_ws_normal_revert = {}
    data_ws_cc_revert = {}
    for id in data_ws:
        pid = data_ws[id]
        for instance in pid:
            ipid = pid[instance]
            row_delete = []
            for row in ipid:
                msg = ''
                rsg = ''
                ripid = ipid[row]
                if ripid[ifcheck] == '???':
                    rsg = 'Info:???????????????{}??????'.format(TU)
                    row_delete.append(row)
                elif ripid[idname] == None and ripid[method] == None:
                    rsg = 'Error:?????????{}???????????????????????????'.format(TU)
                    row_delete.append(row)
                elif ripid[idname] == None:
                    rsg = 'Error:?????????{}??????'.format(TU)
                    row_delete.append(row)
                elif ripid[method] == None:
                    rsg = 'Error:??????{}??????????????????'.format(TU)
                    row_delete.append(row)                
                else:
                    if 'CC' in instance or 'cc' in instance:
                        data_ws_cc_revert.setdefault(id, {})
                        data_ws_cc_revert[id].setdefault(ripid[idname], [])
                        if '?????????' in instance:
                            data_ws_cc_revert[id][ripid[idname]].insert(0, (row, ripid[method], instance))
                        else:
                            data_ws_cc_revert[id][ripid[idname]].append((row, ripid[method], instance))
                    else:
                        data_ws_normal_revert.setdefault(id, {})
                        data_ws_normal_revert[id].setdefault(ripid[idname], [])
                        if '?????????' in instance:
                            data_ws_normal_revert[id][ripid[idname]].insert(0, (row, ripid[method], instance))
                        else:
                            data_ws_normal_revert[id][ripid[idname]].append((row, ripid[method], instance))
                msg = message(msg, rsg)
                mark(ws, 'A', row, msg)
                     

    return data_ws_normal_revert, data_ws_cc_revert


def methodprocess(data_ws, ws, TU):
    for id in data_ws:
        pid = data_ws[id]
        for idname in pid:
            ipid = pid[idname]
            hasbase = False
            if '?????????' in ipid[0][2]:
                checkbase = ipid[0][1]
                for i in range(0, len(ipid)):
                    msg = ''
                    if '?????????' in ipid[i][2]:
                        rsg = 'Info:??????????????????????????????{}??????????????????????????????'.format(idname)
                    else:
                        if ipid[i][1] == checkbase:
                            rsg = 'Info:????????????????????????????????????'
                        else:
                            rsg = 'Error:????????????????????? {}??????????????????????????? {}???????????????'.format(ipid[i][1], checkbase)
                    msg = message(msg, rsg)
                    mark(ws, 'A', ipid[i][0], msg)
            else:
                msg = ''
                rsg = 'Error:????????????????????????{}???????????????????????????????????????'.format(idname)
                msg = message(msg, rsg)
                for i in range(0, len(ipid)):
                    mark(ws, 'A', ipid[i][0], msg)
    return

def methodcheck(data_ws_ori, ws, TU):
    ws.insert_cols(1)
    ws['A1'].value = '????????????????????????'

    data_ws = deepcopy(data_ws_ori)

    data_normal, data_cc = methodpretriage(data_ws, ws, TU)

    methodprocess(data_normal, ws, TU)
    methodprocess(data_cc, ws, TU)
    return

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--cancer", default=r'KN046-301_????????????_20210803.xlsx', help="Please add AE file name")
    parser.add_argument("--bbz", default=r'TUTL|????????????-????????????RECIST 1.1???', help="Please set sheet name of ae")
    parser.add_argument("--fbbz", default=r'TUNTL|????????????-???????????????RECIST 1.1???', help="Please set sheet name of cb")
    parser.add_argument("--xbz", default=r'TUNEWTL|????????????-????????????RECIST 1.1???', help="Please set sheet name of cb")
    parser.add_argument("--recist", default=r'RS|?????????????????????RECIST 1.1???', help="Please set sheet name of cb")

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
        
        bbzcheck(data_ws1, data_ws4, ws1)
        fbbzcheck(data_ws2, data_ws4, ws2)
        xbzcheck(data_ws3, data_ws4, ws3)

        methodcheck(data_ws1, ws1, '?????????')
        methodcheck(data_ws2, ws2, '????????????')

        wb.save(wbsavepath)

    finally:
        wb.close()