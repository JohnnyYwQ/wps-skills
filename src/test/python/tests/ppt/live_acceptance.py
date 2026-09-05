"""Opt-in native PPT acceptance against a fresh disposable presentation."""
import argparse
import json
import os
from pathlib import Path
import sys
import traceback
import uuid
import zipfile
from xml.etree import ElementTree as ET

from wps_skills.client.session_client import ActionFailed, SessionClient
from wps_skills.ppt.demo import create_demo_presentation
from wps_skills.windows.desktop import require_desktop, visible_document_window

REPO=Path(__file__).resolve().parents[5]


def run(output, *, candidate=False):
    desktop=require_desktop()
    output=output.resolve();output.mkdir(parents=True,exist_ok=False)
    path=output/('acceptance-'+uuid.uuid4().hex+'.pptx')
    create_demo_presentation(path)
    alias=output/'alias.pptx';os.link(path,alias)
    report={'status':'running','candidate':candidate,'desktopSessionId':desktop,'steps':[],'checks':[]}
    env=dict(os.environ,PYTHONPATH=str(REPO/'src/main/python'),PYTHONIOENCODING='utf-8')
    def client():
        command=[sys.executable,str(Path(__file__).with_name('session_host_fixture.py'))] if candidate else [sys.executable,'-m','wps_skills.cli.call','--session','--app','ppt']
        return SessionClient(command,application='ppt',env=env,timeout=90)
    def record(): (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    def call(c,name,**params):
        response=c.call({'app':'ppt','action':name},params)
        report['steps'].append({'action':name,'response':response});record()
        return response['data']
    def reject(c,name,code,**params):
        try: call(c,name,**params)
        except ActionFailed as exc:
            assert exc.response['error']['code']==code,exc.response
            report['checks'].append({'rejection':code,'action':name});record()
        else: raise AssertionError('Expected '+code)
    def slide(c,id): return call(c,'getSlideInfo',slideId=id)
    def edit(c,id,name,**params): return call(c,name,slideId=id,expectedToken=slide(c,id)['token'],**params)
    try:
        with client() as c:
            reject(c,'listSlides','SESSION_DOCUMENT_NOT_BOUND')
            reject(c,'openPresentation','DOCUMENT_NOT_FOUND',path=str(output/'missing.pptx'))
            call(c,'openPresentation',path=str(path))
            reject(c,'openPresentation','SESSION_DOCUMENT_ALREADY_BOUND',path=str(path))
            info=call(c,'getPresentationInfo')
            report['window']=visible_document_window(info['window']);assert report['window'],info
            with client() as other:
                reject(other,'openPresentation','DOCUMENT_LEASE_CONFLICT',path=str(path))
                reject(other,'openPresentation','DOCUMENT_LEASE_CONFLICT',path=str(alias))
            structure=call(c,'listSlides');assert structure['slides']==[],structure
            reject(c,'addSlide','INVALID_PARAMS',position=2,expectedToken=structure['token'])
            structure=call(c,'addSlide',position=1,expectedToken=structure['token']);id=structure['slides'][0]['id']
            before=slide(c,id)
            snap=call(c,'addTextBox',slideId=id,expectedToken=before['token'],text='Native PPT\n中文测试',left=40,top=40,width=600,height=100)
            shape=snap['shapes'][0]['id'];assert snap['shapes'][0]['text']=='Native PPT\n中文测试'
            reject(c,'setShapeText','STALE_CONTENT',slideId=id,shapeId=shape,expectedToken=before['token'],text='must not apply')
            edit(c,id,'setShapeText',shapeId=shape,text='Verified PPT\n文本已更新')
            snap=edit(c,id,'formatText',shapeId=shape,format={'latinName':'Arial','eastAsianName':'宋体','size':30,'bold':True,'italic':True,'color':255})
            assert snap['shapes'][0]['font']=={'latinName':'Arial','eastAsianName':'宋体','size':30,'bold':True,'italic':True,'color':255},snap
            edit(c,id,'setShapeGeometry',shapeId=shape,geometry={'left':60,'top':70,'width':640,'height':120})
            snap=edit(c,id,'addShape',kind='rectangle',left=40,top=220,width=200,height=80)
            rectangle=next(s['id'] for s in snap['shapes'] if s['id']!=shape)
            snap=edit(c,id,'addShape',kind='ellipse',left=280,top=220,width=100,height=80)
            assert len(snap['shapes'])==3
            edit(c,id,'deleteShape',shapeId=rectangle)
            reject(c,'getSlideInfo','SLIDE_NOT_FOUND',slideId=99999)
            reject(c,'deleteShape','SHAPE_NOT_FOUND',slideId=id,shapeId=99999,expectedToken=slide(c,id)['token'])
            structure=edit(c,id,'duplicateSlide');copy=structure['slides'][1]['id']
            original=slide(c,id);duplicate=slide(c,copy)
            assert [s['text'] for s in original['shapes']]==[s['text'] for s in duplicate['shapes']]
            structure=call(c,'moveSlide',slideId=copy,position=1,expectedToken=structure['token']);assert structure['slides'][0]['id']==copy
            structure=edit(c,copy,'deleteSlide');assert [s['id'] for s in structure['slides']]==[id]
            before_identity=os.stat(path).st_ino
            saved=call(c,'save');assert saved['documentState']['persistenceState']=='saved'
            after_identity=os.stat(path).st_ino
            report['saveIdentityChanged']=before_identity!=after_identity
            with client() as other:
                reject(other,'openPresentation','DOCUMENT_LEASE_CONFLICT',path=str(path))
                reject(other,'openPresentation','DOCUMENT_LEASE_CONFLICT',path=str(alias))
            edit(c,id,'setShapeText',shapeId=shape,text='Persisted after save\n连续绑定验证成功')
            call(c,'save');assert call(c,'getPresentationInfo')['slideCount']==1
        with client() as reopened:
            call(reopened,'openPresentation',path=str(path))
            snap=slide(reopened,id);assert any(s['text']=='Persisted after save\n连续绑定验证成功' for s in snap['shapes'])
        with zipfile.ZipFile(path) as z:
            texts=[node.text or '' for name in z.namelist() if name.startswith('ppt/slides/slide') and name.endswith('.xml')
                   for node in ET.fromstring(z.read(name)).iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t')]
            assert 'Persisted after save' in texts and '连续绑定验证成功' in texts,texts
        report['checks'].append({'persistedXml':True,'cleanupReacquisition':True,'retainedOldAndNewSaveFences':True})
        report['status']='passed'
    except BaseException as exc:
        report['status']='failed';report['error']=str(exc);report['traceback']=traceback.format_exc();raise
    finally: record()
    return report


def main(argv=None):
    parser=argparse.ArgumentParser();parser.add_argument('--output-dir',type=Path,required=True);parser.add_argument('--candidate',action='store_true')
    args=parser.parse_args(argv);run(args.output_dir,candidate=args.candidate)

if __name__=='__main__': main()
