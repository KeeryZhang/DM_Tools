test = "编号药物:S403703;编号药物:S410496;编号药物:S410502;编号药物:S420679;编号药物:S420954;编号药物:S424979;编号药物:S435358;编号药物:S436543;编号药物:S441286;编号药物:S442701;"
comp = "410502;442701;403703;410496;420679;436543;441286;424979;435358;420954;202007007;122006001"
testlist = test.split(";")

testset = set()
for t in testlist:
    if t == '':
        continue
    name, code = t.split(':')
    if name == '无编号药物':
        continue
    testset.add(code.split('S')[1])

compset = set()
complist = comp.split(';')
for i in complist:
    compset.add(i)

tdc = testset.difference(compset)
cdt = compset.difference(testset)

rsg = 'Error'
rsg += ' {}'.format(' '.join(str(x) for x in cdt))

print(rsg)