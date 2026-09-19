"""Opt-in native regression for whole-row/column COM Range identity.

Generate six independent Task Requests and response expectations. Execute through
an installed Skill in a Windows interactive desktop, never in SSH session 0.
"""
from tests.applications.complex_plan import Case


def cases(root):
    result=[]
    values=[['A','B','C'],[10,11,12],[20,21,22]]
    for axis in ('Rows','Columns'):
        for count in (1,2,3):
            c=Case(axis+str(count),'excel',axis+' insert/delete '+str(count),
                   'Insert the requested axes, verify displaced constants, then delete and restore constants and formula.',root)
            sheet=c.sheet('结构回归')
            c.region('seed','writeRange',sheet,'A1:C3',values=values)
            c.region('formula','setFormulas',sheet,'E5',formulas=[['=SUM(B2:B3)']])
            c.structure('insert','insert'+axis,sheet,start=2,count=count)
            if axis=='Rows':
                address='A1:C'+str(3+count)
                inserted=[values[0]]+[[None]*3 for _ in range(count)]+values[1:]
            else:
                address='A1:'+chr(ord('C')+count)+'3'
                inserted=[[row[0]]+[None]*count+row[1:] for row in values]
            step=c.add('after_insert','readRange',sheet=sheet,address=address)
            c.values(step,inserted)
            c.structure('delete','delete'+axis,sheet,start=2,count=count)
            step=c.add('after_delete','readRange',sheet=sheet,address='A1:C3')
            c.values(step,values)
            step=c.add('formula_restored','readRange',sheet=sheet,address='E5')
            c.values(step,[[32]])
            c.check(step,['cells',0,0,'formula'],'equals','=SUM(B2:B3)')
            c.finish();result.append(c)
    return result
