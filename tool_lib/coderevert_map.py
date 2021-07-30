from outs.Codelist_map import Codelist_map as Codelist_map
import pprint

Coderevert_map_withsymbol = {}
Coderevert_map_nosymbol = {}
AnalyteNamelist = list()
AE_PTlist = list()
Codelist = list(Codelist_map.items())

for i in range(0, len(Codelist)):
    codes = Codelist[i][1]
    for code in codes:
        relist = []
        Coderevert_map_withsymbol.setdefault(code, relist)
        if Codelist[i][0] in Coderevert_map_withsymbol[code]:
            continue
        Coderevert_map_withsymbol[code].append(Codelist[i][0])

       
for AE_PT in Coderevert_map_withsymbol:
    for i in Coderevert_map_withsymbol[AE_PT]:
        AnalyteNamelist.append(i[0])

    for AE in Codelist_map:
        if AE[0] in AnalyteNamelist:
            AE_PTlist.extend([str(x) for x in Codelist_map[AE]])
    AE_PTlist=list(set(AE_PTlist))
    AE_PTlist.sort()
    AE_PTlist = tuple(AE_PTlist)
    Coderevert_map_nosymbol.setdefault(AE_PTlist, AnalyteNamelist)



with open('outs\Coderevert_map.py', 'w') as f:
    f.write('Coderevert_map_withsymbol = ' + pprint.pformat(Coderevert_map_withsymbol) + '\n\n')
    f.write('Coderevert_map_nosymbol =' + pprint.pformat(Coderevert_map_nosymbol) + '\n\n')