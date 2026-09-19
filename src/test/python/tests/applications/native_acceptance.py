"""Opt-in real WPS tests: preplanned JSON -> installed Skill CLI -> Task Response.

Run only in an interactive Windows desktop with --root naming a fresh test area
containing skills/wps-word, skills/wps-excel and skills/wps-ppt. No Agent is used.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import struct
import subprocess
import sys
import time
import zipfile
import zlib


def ref(step, *path):
    return {"$ref": {"step": step, "path": ["data", *path]}}


class Plan:
    def __init__(self, app, path=None):
        self.app = app
        suffix = {"word": "Document", "excel": "Workbook", "ppt": "Presentation"}[app]
        self.request = {"app": app, "document": self.action(('open' if path else 'create') + suffix,
                         **({'path': str(path)} if path else {})), "steps": [], "completion": []}

    def action(self, action_name, **params):
        return {"address": {"app": self.app, "action": action_name}, "params": params}

    def add(self, action_name, **params):
        identifier = 's' + str(len(self.request['steps']) + 1)
        self.request['steps'].append(dict(self.action(action_name, **params), id=identifier))
        return identifier

    def finish(self, output=None, pdf=None, save=False):
        if output:
            self.request['completion'].append(self.action('saveAs', outputPath=str(output), overwritePolicy='failIfExists'))
        elif save:
            self.request['completion'].append(self.action('save'))
        if pdf:
            self.request['completion'].append(self.action('exportPdf', outputPath=str(pdf), overwritePolicy='failIfExists'))
        return self.request


def excel_plan(root):
    p = Plan('excel')
    p.add('getWorkbookInfo')
    sheet = ref(p.add('listWorksheets', offset=0, limit=100), 'worksheets', 0, 'name')
    def edit(action, address, **params):
        s = params.pop('sheet', sheet)
        read = p.add('readRange', sheet=s, address=address)
        return p.add(action, sheet=s, address=address, expectedToken=ref(read, 'token'), **params)
    def structure(action, **params):
        s = params.pop('sheet', sheet)
        read = p.add('getWorksheetInfo', sheet=s)
        return p.add(action, sheet=s, expectedToken=ref(read, 'token'), **params)
    edit('writeRange', 'A1:C4', values=[['名称','数量','金额'],['pear',3,30],['Apple',1,10],['pear',2,20]])
    edit('formatRange', 'A1:C1', format={'bold':True,'italic':True,'fontSize':14,'fontColor':255,'fillColor':65535,'wrapText':True,'horizontalAlignment':'center','verticalAlignment':'center'})
    p.add('findInRange', sheet=sheet,address='A1:C4',text='PEAR',matchCase=False,wholeCell=True,lookIn='values')
    edit('replaceInRange','A2:A4',text='pear',replacement='banana',matchCase=True,wholeCell=False)
    edit('sortRange','A1:C4',column=2,order='ascending',header=True)
    edit('filterRange','A1:C4',column=1,value='Apple')
    edit('clearFilter','A1:C4')
    edit('setRowHeight','A1:C4',height=30)
    edit('setColumnWidth','A1:C4',width=22)
    edit('autoFitColumns','A1:C4')
    src=p.add('readRange',sheet=sheet,address='A1:C4')
    dst=p.add('readRange',sheet=sheet,address='E1:G4')
    p.add('copyRange',sheet=sheet,address='A1:C4',expectedToken=ref(src,'token'),targetSheet=sheet,targetAddress='E1:G4',targetToken=ref(dst,'token'))
    edit('clearRange','E1:G4',mode='all')
    edit('writeRange','A6:B6',values=[['合并标题',None]])
    edit('mergeRange','A6:B6')
    edit('unmergeRange','A6:B6')
    structure('addWorksheet',name='新增')
    structure('renameWorksheet',sheet='新增',name='改名')
    edit('writeRange','A1:B2',sheet='改名',values=[['复制',12],[True,5]])
    structure('copyWorksheet',sheet='改名',name='副本')
    structure('moveWorksheet',sheet='副本',index=1)
    structure('deleteWorksheet',sheet='改名')
    for action in ('insertRows','deleteRows','insertColumns','deleteColumns'):
        structure(action,start=2,count=1)
    edit('setFormulas','C6',formulas=[['=SUM(B2:B4)']])
    edit('calculateRange','C6')
    return p.finish(root/('excel-'+root.parent.name+'.xlsx'),root/('excel-'+root.parent.name+'.pdf'))


def ppt_plan(root):
    p=Plan('ppt')
    p.add('getPresentationInfo')
    listing=p.add('listSlides')
    sid=ref(p.add('addSlide',position=1,expectedToken=ref(listing,'token')),'slides',0,'id')
    def edit(action, **params):
        read=p.add('getSlideInfo',slideId=sid)
        return p.add(action,slideId=sid,expectedToken=ref(read,'token'),**params)
    def style(action, shape, **params):
        read=p.add('getShapeStyle',slideId=sid,shapeId=shape)
        return p.add(action,slideId=sid,shapeId=shape,expectedToken=ref(read,'token'),**params)
    title=ref(edit('addTextBox',text='星河知识库\nHello WPS',left=40,top=40,width=600,height=100),'shapes',0,'id')
    edit('setShapeText',shapeId=title,text='星河知识库\nHello PPT')
    edit('formatText',shapeId=title,format={'size':28,'bold':True,'latinName':'Arial','eastAsianName':'宋体'})
    style('formatShape',title,format={'fillColor':16769230,'lineColor':255,'lineWidth':2})
    style('formatParagraph',title,format={'alignment':'center','spaceBefore':4,'spaceAfter':6,'bulletVisible':False})
    style('setTextBoxLayout',title,layout={'marginLeft':12,'marginRight':12,'verticalAnchor':'middle','wordWrap':True})
    edit('renameShape',shapeId=title,name='项目标题')
    p.add('findText',slideId=sid,text='Hello')
    edit('replaceText',shapeId=title,find='Hello',replacement='你好')
    ids=[]
    for i in range(3):
        step=edit('addShape',kind='rectangle',left=40+i*200,top=200+i*20,width=120,height=40)
        ids.append(ref(step,'shapes',i+1,'id'))
    edit('setShapeGeometry',shapeId=ids[0],geometry={'top':220})
    edit('alignShapes',shapeIds=ids,alignment='top')
    edit('distributeShapes',shapeIds=ids,direction='horizontal')
    edit('setShapeOrder',shapeId=title,position='front')
    read=p.add('getSlideSettings',slideId=sid)
    p.add('setSlideSettings',slideId=sid,expectedToken=ref(read,'token'),settings={'name':'启动页','backgroundColor':16777215})
    read=p.add('getSlideNotes',slideId=sid)
    p.add('setSlideNotes',slideId=sid,expectedToken=ref(read,'token'),text='研发、运营、客服参与。首期四周。')
    edit('addImage',path=str(root/'image.png'),left=650,top=200,width=40,height=40)
    table=edit('addTable',values=[['项目','结果'],['知识库','启动']],left=40,top=320,width=600,height=120)
    tid=ref(table,'shapeId')
    read=p.add('readTable',slideId=sid,shapeId=tid)
    p.add('writeTable',slideId=sid,shapeId=tid,expectedToken=ref(read,'token'),values=[['部门','周期'],['研发/运营/客服','四周']])
    edit('deleteShape',shapeId=ids[2])
    duplicate=edit('duplicateSlide')
    second=ref(duplicate,'slides',1,'id')
    listing=p.add('listSlides')
    p.add('moveSlide',slideId=second,position=1,expectedToken=ref(listing,'token'))
    read=p.add('getSlideInfo',slideId=second)
    p.add('deleteSlide',slideId=second,expectedToken=ref(read,'token'))
    p.add('exportSlideImage',slideId=sid,outputPath=str(root/'slide.png'),overwritePolicy='failIfExists',width=1280,height=720)
    return p.finish(root/('ppt-'+root.parent.name+'.pptx'),root/('ppt-'+root.parent.name+'.pdf'))


def make_png(path):
    def chunk(kind,data):
        return struct.pack('>I',len(data))+kind+data+struct.pack('>I',zlib.crc32(kind+data)&0xffffffff)
    path.write_bytes(b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',2,2,8,2,0,0,0))+
                    chunk(b'IDAT',zlib.compress(b'\x00\xff\x80\x00\x00\x80\xff\x00\x00\x80\xff\xff\x80\x00'))+chunk(b'IEND',b''))


def run(root):
    root=root.resolve()
    out=root/'outputs';out.mkdir(exist_ok=False)
    requests=root/'requests';requests.mkdir(exist_ok=False)
    make_png(out/'image.png')
    env=dict(os.environ,WPS_SKILLS_TASK_DIR=str(root/'receipts'),WPS_TRACE_DIR=str(root/'traces'),PYTHONIOENCODING='utf-8',PYTHONUTF8='1')
    report={'cases':[],'root':str(root),'kind':'direct-json-native-task'}
    def record():
        (root/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    def submit(name,request,expected='succeeded',code=None):
        app=request['app'];path=requests/(name+'.json')
        path.write_text(json.dumps(request,ensure_ascii=False),encoding='utf-8')
        entry=root/'skills'/('wps-'+app)/'scripts'/(app+'.py')
        started=time.perf_counter()
        process=subprocess.run([sys.executable,str(entry),'--app',app,'--task-file',str(path)],env=env,capture_output=True,text=True,encoding='utf-8',timeout=600)
        elapsed=time.perf_counter()-started
        (root/(name+'.stdout')).write_text(process.stdout,encoding='utf-8')
        (root/(name+'.stderr')).write_text(process.stderr,encoding='utf-8')
        lines=[json.loads(line) for line in process.stdout.splitlines() if line.startswith('{')]
        response=next((v for v in reversed(lines) if v.get('type')=='task.response'),None)
        case={'name':name,'app':app,'wallSeconds':elapsed,'exitCode':process.returncode,'response':response,'expected':expected,'expectedCode':code}
        case['passed']=response is not None and response['outcome']==expected and (not code or response['stop']['error']['code']==code)
        if response and response.get('recordPath'):
            status=subprocess.run([sys.executable,str(entry),'--app',app,'--task-status-file',str(path)],env=env,capture_output=True,text=True,encoding='utf-8',timeout=30)
            queried=next((json.loads(line) for line in reversed(status.stdout.splitlines()) if line.startswith('{') and json.loads(line).get('type')=='task.response'),None)
            case['receiptMatches']=queried is not None and all(queried.get(k)==response.get(k) for k in ('taskId','state','outcome','document','steps','completion','cleanup','stop'))
            case['inputConsumed']=not path.exists()
            case['passed'] &= case['receiptMatches'] and case['inputConsumed']
        report['cases'].append(case);record()
        print(name,case['passed'],response.get('stop') if response else process.stderr,flush=True)
        return case
    for app,builder,ext,read_action in [('excel',excel_plan,'xlsx','getWorkbookInfo'),('ppt',ppt_plan,'pptx','getPresentationInfo')]:
        initial=submit(app+'-all-actions',builder(out))
        if not initial['passed']:continue
        path=out/(app+'-'+root.name+'.'+ext)
        with zipfile.ZipFile(path) as z:assert z.testzip() is None
        p=Plan(app,path);p.add(read_action);submit(app+'-open-save',p.finish(save=True))
        # A stale token rejects before mutation and stops the remaining content.
        p=Plan(app,path)
        if app=='excel':
            sh=ref(p.add('listWorksheets',offset=0,limit=100),'worksheets',0,'name')
            token=p.add('readRange',sheet=sh,address='A1')
            p.add('writeRange',sheet=sh,address='A1',expectedToken=ref(token,'token'),values=[['changed']])
            p.add('writeRange',sheet=sh,address='A1',expectedToken=ref(token,'token'),values=[['must not apply']])
            stale_code='STALE_RANGE'
        else:
            sid=ref(p.add('listSlides'),'slides',0,'id');read=p.add('getSlideInfo',slideId=sid)
            shape=ref(read,'shapes',0,'id')
            p.add('setShapeText',slideId=sid,shapeId=shape,expectedToken=ref(read,'token'),text='changed')
            p.add('setShapeText',slideId=sid,shapeId=shape,expectedToken=ref(read,'token'),text='must not apply')
            stale_code='STALE_CONTENT'
        p.add(read_action)
        c=submit(app+'-stale-stop',p.finish(),expected='failed',code=stale_code)
        if c['response']:assert c['response']['steps'][-1]['state']=='not_executed'
        p=Plan(app,path);p.add(read_action)
        c=submit(app+'-existing-changes',p.finish(save=True),expected='failed',code='TASK_EXISTING_CHANGES_CONFIRMATION_REQUIRED')
        if c['response']:assert all(s['state']=='not_executed' for s in c['response']['steps'])
        p=Plan(app,path);p.request['includeExistingChanges']=True
        submit(app+'-authorized-save',p.finish(save=True))
        before=hashlib.sha256(path.read_bytes()).hexdigest()
        p=Plan(app);submit(app+'-output-in-use',p.finish(path),expected='failed',code='OUTPUT_IN_USE')
        assert hashlib.sha256(path.read_bytes()).hexdigest()==before
        occupied=out/('occupied-'+app+'-'+root.name+'.'+ext);shutil.copy2(path,occupied)
        p=Plan(app);submit(app+'-output-exists',p.finish(occupied),expected='failed',code='OUTPUT_ALREADY_EXISTS')
        assert hashlib.sha256(occupied.read_bytes()).hexdigest()==before
        p=Plan(app);submit(app+'-missing-parent',p.finish(out/'missing'/('new.'+ext)),expected='failed',code='OUTPUT_PARENT_NOT_FOUND')
        assert not (out/'missing').exists()
        p=Plan(app,out/('absent.'+ext));p.add(read_action)
        submit(app+'-missing-input',p.finish(),expected='failed',code='DOCUMENT_NOT_FOUND')
    p=Plan('word');p.add('writeContent',anchor={'kind':'documentEnd'},blocks=[{'kind':'paragraph','runs':[{'text':'共享 Task 回归验证'}]}])
    submit('word-shared-regression',p.finish(out/('word-'+root.name+'.docx'),out/('word-'+root.name+'.pdf')))
    applications={c['response']['taskId']:c['app'] for c in report['cases'] if c['response']}
    trace_rows=[]
    for folder in ('requests','tasks','actions'):
        for path in (root/'traces'/folder).glob('*.jsonl'):
            for line in path.read_text(encoding='utf-8').splitlines():
                row=json.loads(line)
                if row.get('taskId') in applications:
                    assert row['app']==applications[row['taskId']], row
                    trace_rows.append(row)
    report['traceApplicationsCorrect']=bool(trace_rows)
    report['passed']=all(c['passed'] for c in report['cases'])
    record()
    return 0 if report['passed'] else 1


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--root',required=True,type=Path)
    raise SystemExit(run(parser.parse_args().root))
