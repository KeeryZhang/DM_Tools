import pprint
from outs.keyslist import keyslist as keysset
keys = ['[Subject]','[AESTDAT_RAW]','[AETCDAT_RAW]','[AEACN]','[AEACN2]','[AEACN3]']

keysset = set(keysset)

for i in keys:
    keysset.add(i)
keyslist = list(keysset)

with open("outs\keyslist.py",'w+') as f:
    f.write('keyslist = ' + pprint.pformat(keyslist))