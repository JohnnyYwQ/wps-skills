"""Live acceptance for common Excel Actions, on one disposable workbook."""
import argparse
import json
import os
from pathlib import Path
import sys
from xml.etree import ElementTree as ET
import zipfile

from wps_skills.client.session_client import ActionFailed, SessionClient
from wps_skills.excel.demo import create_demo_workbook

REPO = Path(__file__).resolve().parents[5]


def run(output, candidate=False):
    output.mkdir(parents=True, exist_ok=False)
    workbook=output/'common-actions.xlsx'
    create_demo_workbook(workbook)
    report={'status':'running','candidate':candidate,'workbook':str(workbook),'steps':[]}
    command=[sys.executable,str(Path(__file__).with_name('session_host_fixture.py'))] if candidate else [sys.executable,'-m','wps_skills.cli.call','--session','--app','excel']
    client=SessionClient(command,application='excel',env=dict(os.environ,PYTHONPATH=str(REPO/'src/main/python'),PYTHONIOENCODING='utf-8'),timeout=120)
    def record():
        (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    def call(action,**params):
        try: response=client.call({'app':'excel','action':action},params)
        except ActionFailed as exc:
            report['steps'].append({'action':action,'params':params,'response':exc.response});record();raise
        report['steps'].append({'action':action,'params':params,'response':response});record()
        return response['data']
    def edit(action,address,**params):
        sheet=params.pop('sheet','演示')
        token=call('readRange',sheet=sheet,address=address)['token']
        return call(action,sheet=sheet,address=address,expectedToken=token,**params)
    def structure(action,sheet='演示',**params):
        token=call('getWorksheetInfo',sheet=sheet)['token']
        return call(action,sheet=sheet,expectedToken=token,**params)
    def rejected(action,code,**params):
        try: call(action,**params)
        except ActionFailed as exc: assert exc.response['error']['code']==code,exc.response
        else: raise AssertionError('Expected '+code)
    try:
        with client:
            call('openWorkbook',path=str(workbook))
            token=call('getWorksheetInfo',sheet='演示')['token']
            rejected('deleteWorksheet','RANGE_UNSUPPORTED',sheet='演示',expectedToken=token)
            edit('writeRange','A1:C4',values=[['名称','数量','金额'],['pear',3,30],['Apple',1,10],['pear',2,20]])
            style={'bold':True,'italic':True,'fontSize':14,'fontColor':255,'fillColor':65535,'wrapText':True,'horizontalAlignment':'center','verticalAlignment':'center'}
            formatted=edit('formatRange','A1:C1',format=style)
            assert all(all(cell[k]==v for k,v in style.items()) for cell in formatted['cells'][0])
            result=call('findInRange',sheet='演示',address='A1:C4',text='PEAR',matchCase=False,wholeCell=True,lookIn='values')
            assert [m['address'] for m in result['matches']]==['A2','A4'],result
            replaced=edit('replaceInRange','A2:A4',text='pear',replacement='banana$',matchCase=True,wholeCell=False)
            assert [r[0]['value'] for r in replaced['cells']]==['banana$','Apple','banana$']
            sorted_=edit('sortRange','A1:C4',column=2,order='ascending',header=True)
            assert [r[1]['value'] for r in sorted_['cells'][1:]]==[1,2,3],sorted_
            filtered=edit('filterRange','A1:C4',column=1,value='Apple')
            assert [r[0]['rowHidden'] for r in filtered['cells']]==[False,False,True,True],filtered
            unfiltered=edit('clearFilter','A1:C4')
            assert not any(r[0]['rowHidden'] for r in unfiltered['cells'])
            edit('setRowHeight','A1:C4',height=30)
            edit('setColumnWidth','A1:C4',width=22)
            fit=edit('autoFitColumns','A1:C4')
            assert all(c['columnWidth']>0 for c in fit['cells'][0])
            source=call('readRange',sheet='演示',address='A1:C4')
            target=call('readRange',sheet='演示',address='E1:G4')
            copied=call('copyRange',sheet='演示',address='A1:C4',expectedToken=source['token'],targetSheet='演示',targetAddress='E1:G4',targetToken=target['token'])
            assert [[c['value'] for c in r] for r in copied['cells']]==[[c['value'] for c in r] for r in source['cells']]
            for mode in ('formats','contents','all'): edit('clearRange','E1:G4',mode=mode)
            edit('writeRange','A6:B6',values=[['合并标题',None]])
            merged=edit('mergeRange','A6:B6')
            assert all(c['merged'] for c in merged['cells'][0])
            unmerged=edit('unmergeRange','A6:B6')
            assert not any(c['merged'] for c in unmerged['cells'][0])
            assert unmerged['cells'][0][0]['value']=='合并标题'
            token=call('readRange',sheet='演示',address='A1:B1')['token']
            rejected('mergeRange','RANGE_UNSUPPORTED',sheet='演示',address='A1:B1',expectedToken=token)
            stale=call('getWorksheetInfo',sheet='演示')['token']
            structure('addWorksheet',name='新增')
            rejected('renameWorksheet','STALE_RANGE',sheet='演示',name='过期',expectedToken=stale)
            structure('renameWorksheet',sheet='新增',name='改名')
            edit('writeRange','A1:B2',sheet='改名',values=[['复制',12],[True,5]])
            structure('copyWorksheet',sheet='改名',name='副本')
            structure('moveWorksheet',sheet='副本',index=1)
            structure('deleteWorksheet',sheet='改名')
            for action in ('insertRows','deleteRows','insertColumns','deleteColumns'):
                structure(action,start=2,count=1)
            check=call('readRange',sheet='演示',address='A1:C4')
            assert [[c['value'] for c in r] for r in check['cells']]==[[c['value'] for c in r] for r in source['cells']]
            # Formula constants must survive literal replacement unchanged.
            edit('setFormulas','C6',formulas=[['=SUM(B2:B4)']])
            formula=edit('replaceInRange','C6',text='SUM',replacement='BAD',matchCase=True,wholeCell=False)
            assert formula['cells'][0][0]['formula']=='=SUM(B2:B4)'
            functions = [
                ('=COUNTIF(B2:B4,">1")',2), ('=COUNTIFS(B2:B4,">1",C2:C4,"<30")',1),
                ('=SUMIF(B2:B4,">1",C2:C4)',50), ('=SUMIFS(C2:C4,B2:B4,">1")',50),
                ('=COUNTA(A1:C4)',12), ('=COUNTBLANK(D1:D4)',4),
                ('=AND(1=1,2>1)',True), ('=OR(1=2,2=2)',True), ('=NOT(FALSE)',True),
                ('=LEFT("abcd",2)','ab'), ('=RIGHT("abcd",2)','cd'), ('=MID("abcd",2,2)','bc'),
                ('=LEN("abcd")',4), ('=TRIM("  a  b  ")','a b'), ('=UPPER("ab")','AB'),
                ('=LOWER("AB")','ab'), ('=CONCATENATE("a","b")','ab'),
                ('=VLOOKUP(2,B2:C4,2,FALSE)',20), ('=HLOOKUP(1,B2:C3,2,FALSE)',2),
                ('=INDEX(C2:C4,2)',20), ('=MATCH(2,B2:B4,0)',2),
                ('=DATE(2024,1,1)',45292), ('=YEAR(DATE(2024,1,1))',2024),
                ('=MONTH(DATE(2024,2,3))',2), ('=DAY(DATE(2024,2,3))',3),
                ('=ROUNDUP(1.21,1)',1.3), ('=ROUNDDOWN(1.29,1)',1.2),
                ('=TEXT(12.5,"0.00")','12.50'), ('=VALUE("12.5")',12.5),
                ('=SUBSTITUTE("aba","a","x")','xbx'), ('=SEARCH("B","abc")',2), ('=FIND("b","abc")',2),
            ]
            formula_address='I1:I'+str(len(functions))
            edit('setFormulas',formula_address,formulas=[[f] for f,v in functions])
            computed=edit('calculateRange',formula_address)
            assert [r[0]['value'] for r in computed['cells']]==[v for f,v in functions],computed
            # Bounds, stale destination and complete merged-area checks happen before mutation.
            token=call('readRange',sheet='演示',address='A1')['token']
            destination=call('readRange',sheet='演示',address='E1')['token']
            edit('writeRange','E1',values=[['changed']])
            rejected('copyRange','STALE_RANGE',sheet='演示',address='A1',expectedToken=token,targetSheet='演示',targetAddress='E1',targetToken=destination)
            stale=call('getWorksheetInfo',sheet='副本')['token']
            edit('writeRange','A1',sheet='副本',values=[['modified']])
            rejected('deleteWorksheet','STALE_RANGE',sheet='副本',expectedToken=stale)
            token=call('getWorksheetInfo',sheet='副本')['token']
            rejected('addWorksheet','INVALID_PARAMS',sheet='副本',name='演示',expectedToken=token)
            edit('mergeRange','E6:F6')
            token=call('readRange',sheet='演示',address='E6')['token']
            rejected('unmergeRange','RANGE_UNSUPPORTED',sheet='演示',address='E6',expectedToken=token)
            edit('unmergeRange','E6:F6')
            call('save')
        assert client.session_outcome['outcome']=='succeeded',client.session_outcome
        ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        with zipfile.ZipFile(workbook) as archive:
            assert archive.testzip() is None
            root=ET.fromstring(archive.read('xl/workbook.xml'))
            names=[s.attrib['name'] for s in root.findall('s:sheets/s:sheet',ns)]
            assert names==['副本','演示'],names
            sheets=[ET.fromstring(archive.read(n)) for n in archive.namelist() if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')]
            assert any(n.text=='SUM(B2:B4)' for sheet in sheets for n in sheet.findall('.//s:f',ns))
        report.update(status='passed',artifactVerified=True)
    except BaseException as exc:
        report.update(status='failed',error=repr(exc),stderr=client.stderr)
        raise
    finally:
        report['sessionOutcome']=client.session_outcome
        record()
    print(output/'report.json')


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--output-dir',required=True,type=Path)
    parser.add_argument('--candidate',action='store_true')
    args=parser.parse_args()
    run(args.output_dir.resolve(),args.candidate)
