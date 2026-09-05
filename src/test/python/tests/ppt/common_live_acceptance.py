"""Opt-in native acceptance of common PPT Actions on a disposable presentation."""
import argparse
import json
import os
from pathlib import Path
import struct
import sys
import traceback
import uuid
import zipfile
import zlib
from xml.etree import ElementTree as ET
from wps_skills.client.session_client import ActionFailed,SessionClient
from wps_skills.ppt.demo import create_demo_presentation
from wps_skills.windows.desktop import require_desktop,visible_document_window

REPO=Path(__file__).resolve().parents[5]


def create_png(path):
    def chunk(kind,data):return struct.pack('>I',len(data))+kind+data+struct.pack('>I',zlib.crc32(kind+data)&0xffffffff)
    path.write_bytes(b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',2,2,8,2,0,0,0))+
        chunk(b'IDAT',zlib.compress(b'\x00\xff\x80\x00\x00\x80\xff\x00\x00\x80\xff\xff\x80\x00'))+chunk(b'IEND',b''))


def run(output,*,candidate=False):
    require_desktop();output=output.resolve();output.mkdir(parents=True,exist_ok=False)
    path=output/('common-'+uuid.uuid4().hex+'.pptx');create_demo_presentation(path)
    image=output/'image.png';create_png(image)
    invalid=output/'invalid.png';invalid.write_text('not a PNG')
    report={'status':'running','candidate':candidate,'steps':[],'checks':[],'presentation':str(path)}
    env=dict(os.environ,PYTHONPATH=str(REPO/'src/main/python'),PYTHONIOENCODING='utf-8')
    def client():
        command=[sys.executable,str(Path(__file__).with_name('session_host_fixture.py'))] if candidate else [sys.executable,'-m','wps_skills.cli.call','--session','--app','ppt']
        return SessionClient(command,application='ppt',env=env,timeout=90)
    def record():(output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    def call(c,action,**params):
        response=c.call({'app':'ppt','action':action},params);report['steps'].append({'action':action,'response':response});record();return response['data']
    def reject(c,action,code,**params):
        try:call(c,action,**params)
        except ActionFailed as exc:
            assert exc.response['error']['code']==code,exc.response
            report['checks'].append({'action':action,'rejection':code});record()
        else:raise AssertionError('Expected '+code)
    def slide(c):return call(c,'getSlideInfo',slideId=sid)
    def edit(c,action,**params):return call(c,action,slideId=sid,expectedToken=slide(c)['token'],**params)
    def style(c,id):return call(c,'getShapeStyle',slideId=sid,shapeId=id)
    def format(c,id,action,**params):return call(c,action,slideId=sid,shapeId=id,expectedToken=style(c,id)['token'],**params)
    try:
        with client() as c:
            call(c,'openPresentation',path=str(path))
            info=call(c,'getPresentationInfo');report['visibleWindow']=visible_document_window(info['window']);assert report['visibleWindow']
            structure=call(c,'listSlides');sid=call(c,'addSlide',position=1,expectedToken=structure['token'])['slides'][0]['id']
            snap=edit(c,'addTextBox',text='😀 Hello WPS\nHello PPT',left=40,top=40,width=700,height=100)
            title=snap['shapes'][0]['id']
            before=style(c,title)
            after=call(c,'formatShape',slideId=sid,shapeId=title,expectedToken=before['token'],format={'fillVisible':True,'fillColor':16769230,'fillTransparency':0.25,'lineVisible':True,'lineColor':255,'lineWidth':2})
            assert abs(after['appearance']['fillTransparency']-0.25)<0.01
            reject(c,'formatShape','STALE_CONTENT',slideId=sid,shapeId=title,expectedToken=before['token'],format={'lineWidth':4})
            format(c,title,'formatParagraph',format={'alignment':'center','spaceBefore':4,'spaceAfter':6,'bulletVisible':True})
            format(c,title,'formatParagraph',format={'alignment':'left','bulletVisible':False})
            layout=format(c,title,'setTextBoxLayout',layout={'marginLeft':12,'marginRight':12,'marginTop':8,'marginBottom':8,'verticalAnchor':'middle','wordWrap':True})
            assert layout['shape']['id']==title
            edit(c,'renameShape',shapeId=title,name='Title')
            matches=call(c,'findText',slideId=sid,text='Hello');assert [m['start'] for m in matches['matches']]==[3,13],matches
            replaced=edit(c,'replaceText',shapeId=title,find='Hello',replacement='你好');assert replaced['replacements']==2
            assert next(s['text'] for s in replaced['shapes'] if s['id']==title)=='😀 你好 WPS\n你好 PPT'
            no_change=edit(c,'replaceText',shapeId=title,find='missing',replacement='unused');assert no_change['replacements']==0
            ids=[]
            for left,top,width in ((50,220,60),(260,260,90),(550,300,120)):
                before=slide(c);after=call(c,'addShape',slideId=sid,expectedToken=before['token'],kind='rectangle',left=left,top=top,width=width,height=40)
                id=next(s['id'] for s in after['shapes'] if s['id'] not in [x['id'] for x in before['shapes']]);ids.append(id)
                format(c,id,'formatShape',format={'fillVisible':True,'fillColor':8795136,'lineVisible':False})
                hidden_line=format(c,id,'formatShape',format={'lineColor':255})
                assert hidden_line['appearance']['lineVisible'] is True
                reject(c,'formatShape','INVALID_PARAMS',slideId=sid,shapeId=id,expectedToken=hidden_line['token'],format={'lineColor':255,'lineVisible':False})
            for alignment in ('left','center','right','top','middle','bottom'):
                # Restore distinct geometry between alignment modes.
                for id,left,top in zip(ids,(50,260,550),(220,260,300)):
                    edit(c,'setShapeGeometry',shapeId=id,geometry={'left':left,'top':top})
                edit(c,'alignShapes',shapeIds=ids,alignment=alignment)
            for id,left,top in zip(ids,(50,260,550),(220,320,440)):edit(c,'setShapeGeometry',shapeId=id,geometry={'left':left,'top':top})
            edit(c,'distributeShapes',shapeIds=ids,direction='horizontal')
            edit(c,'distributeShapes',shapeIds=ids,direction='vertical')
            edit(c,'setShapeOrder',shapeId=title,position='front')
            edit(c,'setShapeOrder',shapeId=title,position='back')
            reject(c,'alignShapes','SHAPE_NOT_FOUND',slideId=sid,shapeIds=[title,99999],alignment='left',expectedToken=slide(c)['token'])
            settings=call(c,'getSlideSettings',slideId=sid)
            call(c,'setSlideSettings',slideId=sid,expectedToken=settings['token'],settings={'name':'常用能力验收','hidden':True,'backgroundColor':16777215})
            reject(c,'setSlideSettings','STALE_CONTENT',slideId=sid,expectedToken=settings['token'],settings={'hidden':False})
            settings=call(c,'getSlideSettings',slideId=sid)
            call(c,'setSlideSettings',slideId=sid,expectedToken=settings['token'],settings={'hidden':False})
            settings=call(c,'getSlideSettings',slideId=sid)
            call(c,'setSlideSettings',slideId=sid,expectedToken=settings['token'],settings={'followMasterBackground':True})
            notes=call(c,'getSlideNotes',slideId=sid)
            call(c,'setSlideNotes',slideId=sid,expectedToken=notes['token'],text='讲者备注\nNative PPT notes 😀')
            reject(c,'setSlideNotes','STALE_CONTENT',slideId=sid,expectedToken=notes['token'],text='must not apply')
            reject(c,'addImage','IMAGE_UNSUPPORTED',slideId=sid,expectedToken=slide(c)['token'],path=str(invalid),left=780,top=40,width=100,height=100)
            reject(c,'addImage','IMAGE_NOT_FOUND',slideId=sid,expectedToken=slide(c)['token'],path=str(output/'missing.png'),left=780,top=40,width=100,height=100)
            edit(c,'addImage',path=str(image),left=780,top=40,width=100,height=100)
            values=[['项目','数值'],['A','42']]
            table=edit(c,'addTable',values=values,left=40,top=350,width=600,height=150);tid=table['shapeId'];assert table['values']==values
            observed=call(c,'readTable',slideId=sid,shapeId=tid)
            reject(c,'writeTable','INVALID_PARAMS',slideId=sid,shapeId=tid,expectedToken=observed['token'],values=[['wrong size']])
            final_values=[['项目','结果'],['中文 😀','已验证']]
            call(c,'writeTable',slideId=sid,shapeId=tid,expectedToken=observed['token'],values=final_values)
            reject(c,'writeTable','STALE_CONTENT',slideId=sid,shapeId=tid,expectedToken=observed['token'],values=values)
            call(c,'save')
            assert call(c,'getSlideNotes',slideId=sid)['text']=='讲者备注\nNative PPT notes 😀'
        with client() as reopened:
            call(reopened,'openPresentation',path=str(path))
            assert call(reopened,'readTable',slideId=sid,shapeId=tid)['values']==final_values
            assert call(reopened,'getSlideSettings',slideId=sid)['name']=='常用能力验收'
        with zipfile.ZipFile(path) as z:
            ns='{http://schemas.openxmlformats.org/drawingml/2006/main}t'
            texts=[n.text or '' for name in z.namelist() if name.startswith('ppt/slides/slide') and name.endswith('.xml') for n in ET.fromstring(z.read(name)).iter(ns)]
            assert '已验证' in texts and any('你好' in t for t in texts)
            notes=[n.text or '' for name in z.namelist() if name.startswith('ppt/notesSlides/notesSlide') and name.endswith('.xml') for n in ET.fromstring(z.read(name)).iter(ns)]
            assert '讲者备注' in notes
            assert any(z.read(name)==image.read_bytes() for name in z.namelist() if name.startswith('ppt/media/')),'PNG was not embedded'
        report['checks'].append({'persistedTableText':True,'persistedNotes':True,'embeddedImage':True})
        report['status']='passed'
    except BaseException as exc:report.update(status='failed',error=str(exc),traceback=traceback.format_exc());raise
    finally:record()
    return report


def main(argv=None):
    parser=argparse.ArgumentParser();parser.add_argument('--output-dir',type=Path,required=True);parser.add_argument('--candidate',action='store_true')
    args=parser.parse_args(argv);run(args.output_dir,candidate=args.candidate)

if __name__=='__main__':main()
