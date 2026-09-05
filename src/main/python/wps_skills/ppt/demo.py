"""Visible WPS demo; all displayed edits use the application Action Session."""
import argparse
from datetime import datetime
import json
import math
from pathlib import Path
import shutil
import time
import uuid

from wps_skills.ppt.skill import open_session
from wps_skills.windows.desktop import require_desktop, visible_document_window

TEMPLATE = Path(__file__).resolve().parents[3] / 'resources/wps_skills/ppt/demo-template.pptx'


def create_demo_presentation(path):
    """Copy the bundled blank fixture; this is not a createPresentation Action."""
    with Path(path).open('xb') as target, TEMPLATE.open('rb') as source:
        shutil.copyfileobj(source, target)


def run_demo(output, *, delay=1.5, session_factory=open_session):
    desktop = require_desktop()
    if not math.isfinite(delay) or not 0 <= delay <= 10:
        raise ValueError('delay must be between 0 and 10 seconds')
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    path = output / ('wps-ppt-' + uuid.uuid4().hex + '.pptx')
    create_demo_presentation(path)
    report = {'status':'running','presentation':str(path),'desktopSessionId':desktop,'steps':[]}
    def call(c, action, params):
        response = c.call({'app':'ppt','action':action},params)
        report['steps'].append(response)
        return response['data']
    def edit(c, slide, action, **params):
        snapshot = call(c,'getSlideInfo',{'slideId':slide})
        return call(c, action,dict(slideId=slide,expectedToken=snapshot['token'],**params))
    def stage(message):
        print(message,flush=True)
        if delay: time.sleep(delay)
    try:
        with session_factory() as c:
            stage('打开 WPS 演示窗口')
            call(c,'openPresentation',{'path':str(path)})
            info = call(c,'getPresentationInfo',{})
            report['visibleWindow'] = visible_document_window(info['window'])
            if not report['visibleWindow']: raise RuntimeError('未能验证所绑定演示文稿的可见窗口。')
            structure = call(c,'listSlides',{})
            structure = call(c,'addSlide',{'position':1,'expectedToken':structure['token']})
            slide = structure['slides'][0]['id']
            stage('添加标题、正文与基础形状')
            snap = edit(c,slide,'addTextBox',text='WPS PPT · 原生自动化',left=60,top=50,width=840,height=90)
            title = next(s['id'] for s in snap['shapes'] if s['text']=='WPS PPT · 原生自动化')
            edit(c,slide,'formatText',shapeId=title,format={'size':36,'bold':True,'color':8795136})
            style = call(c,'getShapeStyle',{'slideId':slide,'shapeId':title})
            call(c,'formatShape',{'slideId':slide,'shapeId':title,'expectedToken':style['token'],
                                 'format':{'fillColor':16769230,'fillVisible':True,'lineVisible':False}})
            snap = edit(c,slide,'addTextBox',text='同一个精确文稿会话\n读取 · 编辑 · 验证 · 保存',left=65,top=180,width=820,height=180)
            body = next(s['id'] for s in snap['shapes'] if s['id']!=title)
            edit(c,slide,'formatText',shapeId=body,format={'size':28,'color':4210752})
            edit(c,slide,'addShape',kind='rectangle',left=65,top=500,width=820,height=8)
            stage('复制幻灯片并修改文本')
            snap = call(c,'getSlideInfo',{'slideId':slide})
            structure=call(c,'duplicateSlide',{'slideId':slide,'expectedToken':snap['token']})
            second=structure['slides'][1]['id']
            snap=call(c,'getSlideInfo',{'slideId':second})
            copied_title=next(s['id'] for s in snap['shapes'] if s['text']=='WPS PPT · 原生自动化')
            edit(c,second,'setShapeText',shapeId=copied_title,text='接入完成 · 动作可读回验证')
            stage('添加原生表格和演讲备注')
            edit(c,second,'addTable',values=[['能力','验证方式'],['形状 / 文字','属性与内容读回'],['表格 / 备注','保存文件独立校验']],
                 left=65,top=350,width=820,height=140)
            notes=call(c,'getSlideNotes',{'slideId':second})
            call(c,'setSlideNotes',{'slideId':second,'expectedToken':notes['token'],
                                   'text':'演示已接入的 33 个 PPT 动作。所有修改通过同一个原生 WPS 会话执行。'})
            stage('显式保存并核对结果')
            report['saved']=call(c,'save',{})
            assert report['saved']['documentState']['persistenceState']=='saved'
            assert len(call(c,'listSlides',{})['slides'])==2
        report['status']='passed'
        print('已保存，WPS 演示窗口保持打开：'+str(path),flush=True)
        return report
    except BaseException as exc:
        report['status']='failed';report['error']=str(exc)
        raise
    finally:
        (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')


def main(argv=None):
    parser=argparse.ArgumentParser(description='Visible native WPS PPT demonstration')
    parser.add_argument('--output-dir',type=Path,default=Path('build')/('ppt-demo-'+datetime.now().strftime('%Y%m%d-%H%M%S')+'-'+uuid.uuid4().hex[:6]))
    parser.add_argument('--delay',type=float,default=1.5)
    args=parser.parse_args(argv)
    run_demo(args.output_dir,delay=args.delay)
    return 0


if __name__=='__main__':
    raise SystemExit(main())
