"""Opt-in native create, first-save, Save As and export acceptance."""
import argparse
import importlib
import json
import os
from pathlib import Path
import sys
import traceback
import uuid
import zipfile
from wps_skills.client.session_client import ActionFailed, SessionClient
from wps_skills.windows.desktop import require_desktop

REPO=Path(__file__).resolve().parents[5]

def run(output, candidate=False):
    desktop=require_desktop()
    output=output.resolve();output.mkdir(parents=True,exist_ok=False)
    report={'status':'running','desktopSessionId':desktop,'candidate':candidate,'steps':[],'checks':[]}
    env=dict(os.environ,PYTHONPATH=str(REPO/'src/main/python'),PYTHONIOENCODING='utf-8')
    def record(): (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    def client(app):
        command=[sys.executable,str(Path(__file__).with_name('session_host_fixture.py')),app] if candidate else [sys.executable,'-m','wps_skills.cli.call','--session','--app',app]
        return SessionClient(command,application=app,env=env,timeout=120)
    def call(c,app,name,**params):
        r=c.call({'app':app,'action':name},params)
        report['steps'].append({'app':app,'action':name,'response':r});record();return r['data']
    def reject(c,app,name,code,**params):
        try: call(c,app,name,**params)
        except ActionFailed as exc:
            assert exc.response['error']['code']==code,exc.response
            assert exc.response['outcome']=='failed',exc.response
            report['checks'].append({'app':app,'action':name,'rejected':code});record()
        else: raise AssertionError('Expected '+code)
    try:
        for app,ext,create,opener in [('word','docx','createDocument','openDocument'),('excel','xlsx','createWorkbook','openWorkbook'),('ppt','pptx','createPresentation','openPresentation')]:
            first=output/(app+'-'+uuid.uuid4().hex+'.'+ext);second=output/(app+'-second-'+uuid.uuid4().hex+'.'+ext)
            existing=output/(app+'-existing.'+ext);existing.write_bytes(b'KEEP EXISTING USER FILE')
            with client(app) as c:
                created=call(c,app,create);assert created['documentState']['persistenceState']=='unsaved',created
                reject(c,app,'save','PERSISTENCE_LOCATOR_REQUIRED')
                reject(c,app,create,'SESSION_DOCUMENT_ALREADY_BOUND')
                reject(c,app,'saveAs','OUTPUT_ALREADY_EXISTS',outputPath=str(existing),overwritePolicy='failIfExists')
                reject(c,app,'saveAs','OUTPUT_PARENT_NOT_FOUND',outputPath=str(output/'missing'/('new.'+ext)),overwritePolicy='failIfExists')
                if app=='excel':
                    sheets=call(c,app,'listWorksheets',offset=0,limit=20);assert len(sheets['worksheets'])==1
                    sheet=sheets['worksheets'][0]['name']
                    def edit(text):
                        snap=call(c,app,'readRange',sheet=sheet,address='A1')
                        return call(c,app,'writeRange',sheet=sheet,address='A1',expectedToken=snap['token'],values=[[text]])
                elif app=='ppt':
                    slides=call(c,app,'listSlides');assert slides['slides']==[]
                    slides=call(c,app,'addSlide',position=1,expectedToken=slides['token']);sid=slides['slides'][0]['id']
                    snap=call(c,app,'getSlideInfo',slideId=sid)
                    snap=call(c,app,'addTextBox',slideId=sid,expectedToken=snap['token'],text='Initial',left=40,top=40,width=600,height=100);shape=snap['shapes'][0]['id']
                    def edit(text):
                        snap=call(c,app,'getSlideInfo',slideId=sid)
                        return call(c,app,'setShapeText',slideId=sid,shapeId=shape,expectedToken=snap['token'],text=text)
                else:
                    def edit(text):
                        return call(c,app,'writeContent',anchor={'kind':'documentEnd'},blocks=[{'kind':'paragraph','runs':[{'text':text}]}])
                edit('Native creation 中文验证')
                if app!='word':
                    pdf=output/(app+'-unsaved.pdf')
                    exported=call(c,app,'exportPdf',outputPath=str(pdf),overwritePolicy='failIfExists')
                    assert exported['documentStateBefore']['persistenceState']=='unsaved'
                    assert pdf.read_bytes().startswith(b'%PDF-')
                    reject(c,app,'exportPdf','OUTPUT_ALREADY_EXISTS',outputPath=str(pdf),overwritePolicy='failIfExists')
                if app=='ppt':
                    png=output/'slide.png';call(c,app,'exportSlideImage',slideId=sid,outputPath=str(png),overwritePolicy='failIfExists',width=1280,height=720)
                    data=png.read_bytes();assert data[:8]==b'\x89PNG\r\n\x1a\n' and int.from_bytes(data[16:20],'big')==1280 and int.from_bytes(data[20:24],'big')==720
                    reject(c,app,'exportSlideImage','SLIDE_NOT_FOUND',slideId=99999,outputPath=str(output/'missing.png'),overwritePolicy='failIfExists',width=1280,height=720)
                saved=call(c,app,'saveAs',outputPath=str(first),overwritePolicy='failIfExists');assert saved['documentState']['persistenceState']=='saved'
                old=first.read_bytes();alias=output/(app+'-alias.'+ext);os.link(first,alias)
                reject(c,app,'saveAs','OUTPUT_MATCHES_BOUND_DOCUMENT',outputPath=str(alias),overwritePolicy='failIfExists')
                with client(app) as other:
                    reject(other,app,opener,'DOCUMENT_LEASE_CONFLICT',path=str(first))
                    reject(other,app,opener,'DOCUMENT_LEASE_CONFLICT',path=str(alias))
                    call(other,app,create)
                    reject(other,app,'save','PERSISTENCE_LOCATOR_REQUIRED')
                edit('After first save')
                call(c,app,'saveAs',outputPath=str(second),overwritePolicy='failIfExists')
                assert first.read_bytes()==old
                with client(app) as other:
                    for path in (first,alias,second): reject(other,app,opener,'DOCUMENT_LEASE_CONFLICT',path=str(path))
                edit('Persisted after Save As')
                saved=call(c,app,'save');assert saved['artifact']['path']==str(second),saved
                if app!='word':
                    exported=call(c,app,'exportPdf',outputPath=str(output/(app+'-saved.pdf')),overwritePolicy='failIfExists')
                    assert exported['documentStateBefore']['persistenceState']=='saved'
            with client(app) as c: call(c,app,opener,path=str(second))
            assert existing.read_bytes()==b'KEEP EXISTING USER FILE'
            with zipfile.ZipFile(second) as z:
                assert z.testzip() is None
                assert any(b'Persisted after Save As' in z.read(n) for n in z.namelist() if n.endswith('.xml'))
            report['checks'].append({'app':app,'persistedXml':True,'retainedOldAndNewLeases':True,'cleanupReacquisition':True});record()
        report['status']='passed'
    except BaseException as exc:
        report.update(status='failed',error=str(exc),traceback=traceback.format_exc());raise
    finally:record()
    return report

def main(argv=None):
    parser=argparse.ArgumentParser();parser.add_argument('--output-dir',type=Path,required=True);parser.add_argument('--candidate',action='store_true');a=parser.parse_args(argv);run(a.output_dir,a.candidate)

if __name__=='__main__': main()
