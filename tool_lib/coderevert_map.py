from outs.Codelist_map import Codelist_map as Codelist_map
import pprint

Coderevert_map = {}

Codelist = list(Codelist_map.items())

for i in range(0, len(Codelist)):
    codes = Codelist[i][1]
    for code in codes:
        relist = []
        Coderevert_map.setdefault(code, relist)
        if Codelist[i][0] in Coderevert_map[code]:
            continue
        Coderevert_map[code].append(Codelist[i][0])

with open('outs\Coderevert_map.py', 'w') as f:
    f.write('Coderevert_map = ' + pprint.pformat(Coderevert_map))