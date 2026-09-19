"""Prepare (never execute) complex Excel/PPT Task requests and response oracles."""
import argparse
import hashlib
import json
from pathlib import Path, PureWindowsPath

from wps_skills.client.applications import compile_request, contracts_for
from wps_skills.client.task_client import _preflight
from tests.applications.native_acceptance import ref, make_png


class Case:
    def __init__(self, identifier, app, title, effect, root, existing=None, setup=False):
        self.id, self.app, self.title, self.effect = identifier, app, title, effect
        self.root, self.setup = root, setup
        self.notes = []
        self.checks = {}
        self.fail = None
        self.shape_counts = {}
        self.initial = ('新建独立空白文档；目标文件均不存在。' if existing is None else
                        '先执行本例 setup 请求生成并保存独立副本；测量请求重新打开该确切路径。初始状态须 saved、非只读。')
        suffix = {'word': 'Document', 'excel': 'Workbook', 'ppt': 'Presentation'}[app]
        self.request = {'app': app, 'document': self.action(
            ('open' if existing else 'create') + suffix,
            **({'path': existing} if existing else {})), 'steps': [], 'completion': []}
        self.check('task_document', ['documentState','persistenceState'], 'equals', 'saved' if existing else 'unsaved')
        self.check('task_document', ['documentState','readOnly'], 'equals', False)
        if existing:self.check('task_document',['artifact','path'],'equals',existing)
        self.existing=existing

    def path(self, suffix):
        return self.root + '/outputs/' + self.id + '-' + self.root.rsplit('/',1)[-1] + '-' + hashlib.sha256(self.root.encode('utf-8')).hexdigest()[:12] + suffix

    def action(self, action_name, **params):
        return {'address':{'app':self.app,'action':action_name},'params':params}

    def add(self, identifier, action_name, **params):
        if any(s['id']==identifier for s in self.request['steps']):raise ValueError(identifier)
        self.request['steps'].append(dict(self.action(action_name, **params),id=identifier))
        return identifier

    def check(self, step, path, op, value, tolerance=None):
        rule={'path':path,'op':op,'value':value}
        if tolerance is not None:rule['tolerance']=tolerance
        self.checks.setdefault(step,[]).append(rule)

    def finish(self, pdf=False, save=False):
        path=self.existing if save else self.path({'word': '.docx', 'excel': '.xlsx', 'ppt': '.pptx'}[self.app])
        self.request['completion'].append(self.action('save') if save else self.action('saveAs',outputPath=path,overwritePolicy='failIfExists'))
        self.check('task_save',['artifact','path'],'equals',path)
        self.check('task_save',['artifact','format'],'equals',{'word': 'docx', 'excel': 'xlsx', 'ppt': 'pptx'}[self.app])
        self.check('task_save',['artifact','sizeBytes'],'greater_than',0)
        self.check('task_save',['documentState','persistenceState'],'equals','saved')
        if pdf:
            self.request['completion'].append(self.action('exportPdf',outputPath=self.path('.pdf'),overwritePolicy='failIfExists'))
            self.check('task_pdf',['artifact','path'],'equals',self.path('.pdf'))
            self.check('task_pdf',['artifact','format'],'equals','pdf')
            self.check('task_pdf',['artifact','sizeBytes'],'greater_than',0)
            self.check('task_pdf',['documentStateBefore'],'equals',ref('task_pdf','documentStateAfter'))
            self.check('task_pdf',['documentStateAfter','persistenceState'],'equals','saved')
        return self

    def expect_failure(self, step, code):
        self.fail=(step,code)

    def expectation(self):
        plan=compile_request(self.request,self.app)
        _preflight(plan,contracts_for(self.app))
        records=[];stopped=False
        for s in plan['steps']:
            failed=self.fail is not None and s['id']==self.fail[0]
            state='not_executed' if stopped else 'failed' if failed else 'succeeded'
            record={'id':s['id'],'address':s['address'],'state':state,
                    'responseOutcome':None if stopped else state,
                    'dataAssertions':self.checks.get(s['id'],[]) if state=='succeeded' else []}
            if failed:record['errorCode']=self.fail[1];stopped=True
            records.append(record)
        if self.fail and not stopped:raise ValueError('missing expected failure')
        return {'caseId':self.id,'kind':'negative' if self.fail else 'positive','setupOnly':self.setup,
                'task':{'app':self.app,'outcome':'failed' if self.fail else 'succeeded',
                        'state':'stopped' if self.fail else 'completed',
                        'stop':{'stepId':self.fail[0],'phase':'execution','error':{'code':self.fail[1]}} if self.fail else None,
                        'cleanupOutcome':'succeeded','taskFileState':'removed','receiptMatches':True},
                'actions':records,'humanExpected':self.effect,'initialConditions':self.initial,
                'notes':self.notes,'requestFile':self.root+'/requests/'+self.id+'.json'}

    def sheet(self, name='明细'):
        s=self.add('sheets','listWorksheets',offset=0,limit=100)
        old=ref(s,'worksheets',0,'name')
        self.structure('name_sheet','renameWorksheet',old,name=name)
        return name

    def structure(self, identifier, action, sheet, **params):
        read=self.add(identifier+'_token','getWorksheetInfo',sheet=sheet)
        s=self.add(identifier,action,sheet=sheet,expectedToken=ref(read,'token'),**params)
        if action=='deleteWorksheet':self.check(s,['deleted'],'equals',sheet)
        else:self.check(s,['name'],'equals',params.get('name',sheet))
        return s

    def region(self, identifier, action, sheet, address, **params):
        read=self.add(identifier+'_token','readRange',sheet=sheet,address=address)
        s=self.add(identifier,action,sheet=sheet,address=address,expectedToken=ref(read,'token'),**params)
        self.check(s,['sheet'],'equals',sheet);self.check(s,['address'],'equals',address)
        if action=='writeRange':self.values(s,params['values'])
        if action=='setFormulas':self.check(s,['cells','*','*','formula'],'equals',params['formulas'])
        if action=='formatRange':
            for k,v in params['format'].items():self.check(s,['cells','*','*',k],'all_equals',v)
        return s

    def values(self, step, expected):
        self.check(step,['cells','*','*','value'],'equals',expected)
        self.check(step,['cells','*','*','errorCode'],'all_equals',None)

    def copy(self, identifier, source, address, target, dest):
        a=self.add(identifier+'_source','readRange',sheet=source,address=address)
        b=self.add(identifier+'_target','readRange',sheet=target,address=dest)
        s=self.add(identifier,'copyRange',sheet=source,address=address,expectedToken=ref(a,'token'),targetSheet=target,targetAddress=dest,targetToken=ref(b,'token'))
        self.check(s,['sheet'],'equals',target);self.check(s,['address'],'equals',dest)
        self.check(s,['cells','*','*','formula'],'all_equals',None)
        # Expected-value selectors are separate from executable Task references.
        self.check(s,['cells','*','*','value'],'equals',{'$response':{'step':a,'path':['cells','*','*','value']}})
        return s

    def slide(self, identifier, position):
        r=self.add(identifier+'_list','listSlides')
        s=self.add(identifier,'addSlide',position=position,expectedToken=ref(r,'token'))
        sid=ref(s,'slides',position-1,'id');self.shape_counts[identifier]=0
        self.check(s,['slides',position-1,'index'],'equals',position)
        return sid

    def edit(self, identifier, action, sid, **params):
        r=self.add(identifier+'_token','getSlideInfo',slideId=sid)
        s=self.add(identifier,action,slideId=sid,expectedToken=ref(r,'token'),**params)
        if action not in {'duplicateSlide','deleteSlide','addTable'}:self.check(s,['slide','id'],'equals',sid)
        return s

    def text(self, identifier, sid, text, index=0, left=40,top=35,width=630,height=65):
        s=self.edit(identifier,'addTextBox',sid,text=text,left=left,top=top,width=width,height=height)
        shape=ref(s,'shapes',index,'id')
        self.shape(s,shape,{'text':text,'left':left,'top':top,'width':width,'height':height})
        return shape

    def shape(self, step, sid, fields):
        for k,v in fields.items():
            self.check(step,['shapes',{'field':'id','value':sid},*k.split('.')],'close' if k in {'left','top','width','height','rotation','font.size'} else 'equals',v,0.1 if k in {'left','top','width','height','rotation','font.size'} else None)

    def style(self, identifier, action, sid, shape, **params):
        r=self.add(identifier+'_token','getShapeStyle',slideId=sid,shapeId=shape)
        s=self.add(identifier,action,slideId=sid,shapeId=shape,expectedToken=ref(r,'token'),**params)
        self.check(s,['shape','id'],'equals',shape)
        if action=='formatShape':
            for k,v in params['format'].items():self.check(s,['appearance',k],'close' if k in {'lineWidth','fillTransparency'} else 'equals',v,0.01 if k in {'lineWidth','fillTransparency'} else None)
        elif action=='setTextBoxLayout':
            for k,v in params['layout'].items():
                if k=='verticalAnchor':v={'top':1,'middle':3,'bottom':4}[v]
                self.check(s,['textBox',k],'equals' if isinstance(v,bool) or k=='verticalAnchor' else 'close',v,0.1 if not isinstance(v,bool) and k!='verticalAnchor' else None)
        elif action=='formatParagraph':
            for k,v in params['format'].items():
                if k=='alignment':v={'left':1,'center':2,'right':3,'justify':4}[v]
                self.check(s,['paragraphs','*',k],'all_equals',v)
        return s

    def font(self, identifier, sid, shape, **patch):
        s=self.edit(identifier,'formatText',sid,shapeId=shape,format=patch)
        self.shape(s,shape,{'font.'+k:v for k,v in patch.items()});return s

    def notes_for(self, identifier, sid, text):
        r=self.add(identifier+'_token','getSlideNotes',slideId=sid)
        s=self.add(identifier,'setSlideNotes',slideId=sid,expectedToken=ref(r,'token'),text=text)
        self.check(s,['text'],'equals',text);return s

    def settings(self, identifier, sid, **patch):
        r=self.add(identifier+'_token','getSlideSettings',slideId=sid)
        s=self.add(identifier,'setSlideSettings',slideId=sid,expectedToken=ref(r,'token'),settings=patch)
        for k,v in patch.items():self.check(s,[k],'equals',v)
        return s

    def png(self, identifier,sid):
        path=self.path('-'+identifier+'.png')
        s=self.add(identifier,'exportSlideImage',slideId=sid,outputPath=path,overwritePolicy='failIfExists',width=1280,height=720)
        self.check(s,['artifact','path'],'equals',path);self.check(s,['artifact','sizeBytes'],'greater_than',0)
        self.check(s,['documentStateBefore'],'equals',ref(s,'documentStateAfter'))
        return s


def excel_cases(root):
    cases=[];setup=[]
    def new(n,title,effect,existing=None):
        c=Case('X%02d'%n,'excel',title,effect,root,existing);cases.append(c);return c
    c=new(1,'三部门月度报表与汇总','研发、运营、客服各一张明细表，另有汇总表；汇总值分别为 78、156、234，合计 468。表头加粗、数量使用整数格式。')
    sheet=c.sheet('研发')
    totals=[]
    for i,name in enumerate(('研发','运营','客服')):
        if i:c.structure('add_'+str(i),'addWorksheet','研发',name=name)
        values=[['任务','负责人','数量','状态']]+[['任务%02d'%j,name+'同事',j*(i+1),'完成' if j%2 else '进行中'] for j in range(1,13)]
        c.region('write_'+str(i),'writeRange',name,'A1:D13',values=values)
        c.region('head_'+str(i),'formatRange',name,'A1:D1',format={'bold':True,'fillColor':15921906})
        c.region('sum_'+str(i),'setFormulas',name,'C15',formulas=[['=SUM(C2:C13)']])
        calc=c.region('calc_'+str(i),'calculateRange',name,'C15');c.values(calc,[[78*(i+1)]])
        totals.append(ref(calc,'cells',0,0,'value'))
    c.structure('summary','addWorksheet','研发',name='汇总')
    c.region('summary_values','writeRange','汇总','A1:B4',values=[['部门','数量'],*[[n,v] for n,v in zip(('研发','运营','客服'),totals)]])
    c.region('total_formula','setFormulas','汇总','B6',formulas=[['=SUM(B2:B4)']])
    s=c.region('total_calc','calculateRange','汇总','B6');c.values(s,[[468]])
    c.finish()

    c=new(2,'公式链、条件判断与受控错误','8 条明细计算金额、折后金额与达标标记；汇总 360、324、4；IFERROR 捕获除零返回 0，未捕获除零单元格单独返回错误码而非普通数值。')
    sh=c.sheet();data=[['数量','单价','金额','折后','标记']]+[[i,10,None,None,None] for i in range(1,9)]
    c.region('data','writeRange',sh,'A1:E9',values=data)
    fs=[['=A%d*B%d'%(i,i),'=ROUND(C%d*0.9,2)'%i,'=IF(D%d>=40,"达标","未达标")'%i] for i in range(2,10)]
    c.region('formulas','setFormulas',sh,'C2:E9',formulas=fs)
    s=c.region('calculate','calculateRange',sh,'C2:E9');c.values(s,[[i*10,i*9,'达标' if i*9>=40 else '未达标'] for i in range(1,9)])
    c.region('summary_formula','setFormulas',sh,'G2:G6',formulas=[['=SUM(C2:C9)'],['=SUM(D2:D9)'],['=COUNTIF(E2:E9,"达标")'],['=IFERROR(1/0,0)'],['=1/0']])
    s=c.region('summary_calc','calculateRange',sh,'G2:G6')
    c.check(s,['cells','*',0,'value'],'equals',[360,324,4,0,None]);c.check(s,['cells',4,0,'errorCode'],'not_null',None)
    c.notes.append('除零错误作为 calculateRange 的单元格观察值验证；Task 仍预期成功，错误码必须独立于 value。')
    c.finish()

    c=new(3,'混合类型与分区格式','保留中文、英文、布尔值、空单元格、前导零文本、看似公式的字面文本和 emoji；金额两位小数、百分比格式、长文本换行。')
    sh=c.sheet('数据类型')
    data=[['编号','金额','比例','启用','备注'],['001',12.5,0.25,True,'中文 English'],['002',0,0,False,'第一行\n第二行'],['003',-3.25,1,None,'=SUM(A1:A2)'],['004',999.99,0.875,True,'😀 测试'],['005',None,None,False,'尾部文本']]
    c.region('mixed','writeRange',sh,'A1:E6',values=data)
    for ident,addr,patch in [('head','A1:E1',{'bold':True,'fontSize':14,'horizontalAlignment':'center'}),('money','B2:B6',{'numberFormat':'0.00'}),('ratio','C2:C6',{'numberFormat':'0.0%'}),('notes','E2:E6',{'wrapText':True,'italic':True,'fontSize':11})]:c.region(ident,'formatRange',sh,addr,format=patch)
    s=c.region('row_height','setRowHeight',sh,'A1:E6',height=32);c.values(s,data);c.check(s,['cells','*','*','rowHeight'],'all_close',32,0.8)
    c.finish()

    c=new(4,'排序、筛选与恢复显示','12 条工单按金额降序排列，整行关系保持；筛选只显示研发，再清除筛选，全部明细重新可见。')
    sh=c.sheet('工单');rows=[['单号','部门','金额']]+[['T%02d'%i,('研发','运营','客服')[(i-1)%3],i*10] for i in range(1,13)]
    c.region('data','writeRange',sh,'A1:C13',values=rows)
    s=c.region('sort','sortRange',sh,'A1:C13',column=3,order='descending',header=True);ordered=[rows[0]]+list(reversed(rows[1:]));c.values(s,ordered)
    s=c.region('filter','filterRange',sh,'A1:C13',column=2,value='研发');c.check(s,['cells','*',0,'rowHidden'],'equals',[False]+[r[1]!='研发' for r in ordered[1:]])
    s=c.region('clear_filter','clearFilter',sh,'A1:C13');c.check(s,['cells','*',0,'rowHidden'],'all_equals',False);c.values(s,ordered)
    c.region('header','formatRange',sh,'A1:C1',format={'bold':True,'fillColor':13434879});c.finish()

    c=new(5,'跨表复制数值、公式值固化与源数据保留','源表中公式计算出 20、60、120、200、300、420；复制到交付表后这些格为数值且不含公式。清空一份临时副本，不影响交付区域。')
    sh=c.sheet('源数据');c.structure('target','addWorksheet',sh,name='交付')
    c.region('data','writeRange',sh,'A1:C7',values=[['数量','单价','金额']]+[[i,i*10,None] for i in range(1,7)])
    c.region('f','setFormulas',sh,'C2:C7',formulas=[['=A%d*B%d'%(r,r)] for r in range(2,8)])
    s=c.region('calc','calculateRange',sh,'C2:C7');c.values(s,[[i*i*10] for i in range(1,7)])
    c.copy('delivery',sh,'A1:C7','交付','A1:C7');c.copy('temporary',sh,'A1:C7','交付','E1:G7')
    s=c.region('clear','clearRange','交付','E1:G7',mode='contents');c.check(s,['cells','*','*','value'],'all_equals',None)
    c.region('delivery_style','formatRange','交付','A1:C7',format={'numberFormat':'0.00'})
    c.notes.append('复制为 values，不把复制结果误判为公式复制。')
    c.effect=c.effect.replace('20、60、120、200、300、420','10、40、90、160、250、360')
    c.finish()

    c=new(6,'工作表复制、重命名、移动与删除','最终有“交付版、模板、归档”三张表；交付版来自模板复制并有独立标题，临时表已删除，模板标题仍为原值。')
    sh=c.sheet('模板');c.region('template','writeRange',sh,'A1:C4',values=[['模板原值','数量','说明'],['A',10,'保留'],['B',20,'保留'],['C',30,'保留']])
    c.structure('copy','copyWorksheet',sh,name='草稿');c.structure('rename','renameWorksheet','草稿',name='交付版');c.structure('move','moveWorksheet','交付版',index=1)
    c.structure('archive','addWorksheet',sh,name='归档');c.structure('temp','addWorksheet','归档',name='临时');c.structure('delete','deleteWorksheet','临时')
    c.region('update','writeRange','交付版','A1',values=[['正式交付']])
    r=c.region('template_style','formatRange','模板','A1:C1',format={'bold':True});c.check(r,['cells',0,0,'value'],'equals','模板原值')
    # Listing supplies the destination sheet name for the final title formatting.
    s=c.add('final_order','listWorksheets',offset=0,limit=100);c.check(s,['worksheets','*','name'],'equals',['交付版','模板','归档'])
    c.region('final_style','formatRange',ref(s,'worksheets',0,'name'),'A1:C1',format={'bold':True,'fontSize':16});c.finish()

    c=new(7,'行列插删后的数据与公式调整','在 6 条明细中插入并移除两行、插入并移除一列；最终数据顺序恢复原样，合计公式恢复 =SUM(B2:B7)，结果 210。')
    sh=c.sheet('结构调整');values=[['编号','数量','说明']]+[['R%d'%i,i*10,'原始'] for i in range(1,7)]
    c.region('base','writeRange',sh,'A1:C7',values=values);c.region('sum','setFormulas',sh,'B9',formulas=[['=SUM(B2:B7)']])
    c.structure('insert_rows','insertRows',sh,start=3,count=2);c.region('new_rows','writeRange',sh,'A3:C4',values=[['临时1',1,'待删'],['临时2',2,'待删']])
    c.structure('delete_rows','deleteRows',sh,start=3,count=2);c.structure('insert_col','insertColumns',sh,start=2,count=1)
    c.region('temporary_col','writeRange',sh,'B1:B7',values=[['临时列']]+[[i] for i in range(1,7)])
    c.structure('delete_col','deleteColumns',sh,start=2,count=1)
    s=c.region('style','formatRange',sh,'A1:C7',format={'fontSize':11});c.values(s,values)
    s=c.region('calc','calculateRange',sh,'B9');c.values(s,[[210]]);c.check(s,['cells',0,0,'formula'],'equals','=SUM(B2:B7)');c.finish()

    c=new(8,'分组标题合并、取消合并与重排','三组四列表头只保留左上角标题；第二组取消合并后写入四个列名，第一和第三组继续合并，正文数据不受影响。')
    sh=c.sheet('分组报表')
    for i,row in enumerate((1,5,9)):
        c.region('title%d'%i,'writeRange',sh,'A%d:D%d'%(row,row),values=[['分组%d'%(i+1),None,None,None]])
        c.region('style%d'%i,'formatRange',sh,'A%d:D%d'%(row,row),format={'bold':True,'horizontalAlignment':'center'})
        s=c.region('merge%d'%i,'mergeRange',sh,'A%d:D%d'%(row,row));c.check(s,['cells','*','*','merged'],'all_equals',True)
        c.region('body%d'%i,'writeRange',sh,'A%d:D%d'%(row+1,row+2),values=[['A',10,20,30],['B',40,50,60]])
    s=c.region('unmerge','unmergeRange',sh,'A5:D5');c.check(s,['cells','*','*','merged'],'all_equals',False)
    c.region('new_head','writeRange',sh,'A5:D5',values=[['项目','一月','二月','三月']]);c.finish()

    c=new(9,'查找替换与公式保护','把“待确认”文字常量替换为“已确认”；含这些字的公式字符串保持原样；查找分别返回值匹配和公式匹配的确切单元格。')
    sh=c.sheet('确认清单');data=[['状态','备注'],['待确认','待确认-研发'],['已完成','待确认-运营'],['待确认','无变更'],['进行中','结束']]
    c.region('data','writeRange',sh,'A1:B5',values=data)
    c.region('f','setFormulas',sh,'C2:C3',formulas=[['="待确认"'],['=IF(1=1,"待确认","无")']])
    s=c.add('find_values','findInRange',sheet=sh,address='A1:B5',text='待确认',matchCase=True,wholeCell=False,lookIn='values')
    c.check(s,['matches','*','address'],'equals',['A2','B2','B3','A4'])
    s=c.region('replace','replaceInRange',sh,'A1:C5',text='待确认',replacement='已确认',matchCase=True,wholeCell=False)
    c.check(s,['cells',1,0,'value'],'equals','已确认');c.check(s,['cells',1,1,'value'],'equals','已确认-研发')
    c.check(s,['cells',1,2,'formula'],'equals','="待确认"');c.check(s,['cells',2,2,'formula'],'equals','=IF(1=1,"待确认","无")')
    s=c.add('find_formulas','findInRange',sheet=sh,address='C2:C3',text='待确认',matchCase=True,wholeCell=False,lookIn='formulas')
    c.check(s,['matches','*','address'],'equals',['C2','C3']);c.finish()

    seed=Case('X10-setup','excel','准备 X10 已有文件','创建并保存“已有数据”工作簿副本。',root,setup=True);sh=seed.sheet('已有数据')
    seed.region('seed','writeRange',sh,'A1:C5',values=[['项目','数量','状态'],['A',10,'待处理'],['B',20,'待处理'],['C',30,'完成'],['D',40,'完成']]);seed.finish();setup.append(seed)
    c=new(10,'已有文件编辑与原路径保存','打开本例预先保存的副本；将 A/B 数量改为 15/25，待处理改成处理中；复制一张归档表，原路径保存并导出 PDF。',seed.path('.xlsx'))
    s=c.add('locate','listWorksheets',offset=0,limit=100);sh=ref(s,'worksheets',0,'name')
    c.region('amounts','writeRange',sh,'B2:B3',values=[[15],[25]])
    c.region('status','replaceInRange',sh,'C2:C5',text='待处理',replacement='处理中',matchCase=True,wholeCell=True)
    c.structure('archive','copyWorksheet',sh,name='本轮归档');c.region('format','formatRange','本轮归档','A1:C5',format={'fontSize':12})
    c.finish(pdf=True,save=True)

    c=new(11,'较大明细、公式列与双格式交付','40 条、8 列业务明细；第七列为数量×单价，第八列为折后金额。汇总金额 8200、折后 7380，交付 XLSX 与 PDF。')
    sh=c.sheet('月度明细');data=[['序号','部门','项目','数量','单价','折扣','金额','折后']]+[[i,('研发','运营','客服')[(i-1)%3],'条目%02d'%i,i,10,0.9,None,None] for i in range(1,41)]
    c.region('bulk_data','writeRange',sh,'A1:H41',values=data)
    c.region('bulk_formula','setFormulas',sh,'G2:H41',formulas=[['=D%d*E%d'%(i,i),'=ROUND(G%d*F%d,2)'%(i,i)] for i in range(2,42)])
    s=c.region('calc','calculateRange',sh,'G2:H41');c.values(s,[[i*10,i*9] for i in range(1,41)])
    c.region('totals','setFormulas',sh,'G43:H43',formulas=[['=SUM(G2:G41)','=SUM(H2:H41)']]);s=c.region('totals_calc','calculateRange',sh,'G43:H43');c.values(s,[[8200,7380]])
    c.region('header','formatRange',sh,'A1:H1',format={'bold':True,'fontSize':12,'fillColor':15921906})
    c.region('width','setColumnWidth',sh,'A1:H1',width=12)
    c.notes.append('最大写入 328 格，位于 1000 格限制以内；沿用 60 秒 Action deadline，不因失败自动加时重试。PDF 使用当前原生打印设置，不预设页数。')
    c.finish(pdf=True)

    c=new(12,'复杂链中的过期凭据与停止','写入 8 行清单并修改一个区域后，故意用修改前 token 再写；第二次写被拒绝。前面的内容留在未保存文档中，后续格式、保存、PDF 均不执行。')
    sh=c.sheet('停止验证');c.region('data','writeRange',sh,'A1:C9',values=[['编号','金额','状态']]+[['A%d'%i,i*10,'原始'] for i in range(1,9)])
    r=c.add('old_token','readRange',sheet=sh,address='B2:B5')
    s=c.add('first_write','writeRange',sheet=sh,address='B2:B5',expectedToken=ref(r,'token'),values=[[101],[102],[103],[104]]);c.values(s,[[101],[102],[103],[104]])
    c.add('stale_write','writeRange',sheet=sh,address='B2:B5',expectedToken=ref(r,'token'),values=[[999],[999],[999],[999]])
    c.region('must_not_format','formatRange',sh,'A1:C1',format={'bold':True});c.finish(pdf=True);c.expect_failure('stale_write','STALE_RANGE')
    return cases,setup


def ppt_cases(root):
    cases=[];setup=[]
    def new(n,title,effect,existing=None):
        c=Case('P%02d'%n,'ppt',title,effect,root,existing);cases.append(c);return c
    c=new(1,'六页项目启动演示','六页依次为项目概览、目标、分工、里程碑、风险、下一步；每页有标题、三项正文及备注，标题 28pt 加粗，正文 18pt。')
    slides=[]
    for i,title in enumerate(('项目概览','目标','分工','里程碑','风险','下一步'),1):
        sid=c.slide('page%d'%i,i);slides.append(sid)
        head=c.text('title%d'%i,sid,title)
        body=c.text('body%d'%i,sid,'星河知识库\n研发、运营、客服共同参与\n首期四周交付',index=1,top=120,height=220)
        c.font('head_font%d'%i,sid,head,size=28,bold=True)
        c.font('body_font%d'%i,sid,body,size=18)
        c.notes_for('notes%d'%i,sid,'第%d页：%s，说明负责人及验收要求。'%(i,title))
    c.check('page6',['slides'],'length',6);c.finish(pdf=True)

    c=new(2,'中西文字体、字号与连续改写','同一页四个文本框分别使用不同字号和强调样式；中文宋体、英文 Arial，最后改写第二框并重新设置 20pt、蓝色、非斜体。')
    sid=c.slide('page',1)
    for i in range(4):
        shape=c.text('box%d'%i,sid,'星河 Knowledge Base %d\n中文与 English 混排'%(i+1),index=i,top=25+i*110,height=95)
        c.font('font%d'%i,sid,shape,latinName='Arial',eastAsianName='宋体',size=16+i*4,bold=i%2==0,italic=i%2==1,color=255 if i==0 else 0)
        if i==1:second=shape
    s=c.edit('rewrite','setShapeText',sid,shapeId=second,text='第二框已更新\nFinal content');c.shape(s,second,{'text':'第二框已更新\nFinal content'})
    c.font('rewrite_font',sid,second,latinName='Arial',eastAsianName='宋体',size=20,color=16711680,italic=False)
    c.finish()

    c=new(3,'多段文字与文本框布局','三列各含四段内容；分别左、中、右对齐，段前 4pt、段后 8pt，第二列带项目符号。文本框边距明确、换行开启、顶端对齐。')
    sid=c.slide('page',1)
    for i,align in enumerate(('left','center','right')):
        shape=c.text('box%d'%i,sid,'栏目%d\n第一条说明\n第二条说明\n最后一段'%(i+1),index=i,left=30+i*230,top=70,width=210,height=350)
        c.font('font%d'%i,sid,shape,size=18)
        c.style('para%d'%i,'formatParagraph',sid,shape,format={'alignment':align,'spaceBefore':4,'spaceAfter':8,'bulletVisible':i==1})
        c.style('layout%d'%i,'setTextBoxLayout',sid,shape,layout={'marginLeft':12,'marginRight':12,'marginTop':10,'marginBottom':10,'verticalAnchor':'top','wordWrap':True})
    c.finish()

    c=new(4,'多形状对齐与等间距分布','上方三矩形顶边对齐、水平等间距；右下三矩形左边对齐、垂直等间距。六个形状均保留原尺寸。')
    sid=c.slide('page',1);groups=[]
    for group in range(2):
        ids=[]
        for i in range(3):
            left=(40,250,560)[i] if group==0 else (370,390,410)[i]
            top=(50,70,90)[i] if group==0 else (170,270,420)[i]
            step=c.edit('rect%d_%d'%(group,i),'addShape',sid,kind='rectangle',left=left,top=top,width=100,height=40)
            ids.append(ref(step,'shapes',group*3+i,'id'))
        groups.append(ids)
    c.edit('align_top','alignShapes',sid,shapeIds=groups[0],alignment='top')
    s=c.edit('horizontal','distributeShapes',sid,shapeIds=groups[0],direction='horizontal')
    for shape,left in zip(groups[0],(40,300,560)):c.shape(s,shape,{'left':left,'top':50,'width':100,'height':40})
    c.edit('align_left','alignShapes',sid,shapeIds=groups[1],alignment='left')
    s=c.edit('vertical','distributeShapes',sid,shapeIds=groups[1],direction='vertical')
    for shape,top in zip(groups[1],(170,295,420)):c.shape(s,shape,{'left':370,'top':top,'width':100,'height':40})
    c.finish()

    c=new(5,'层叠形状、外观与几何调整','背景矩形、说明文字和强调框形成三层；背景置底，文字置顶；强调框移到指定位置、改变尺寸和边框，最后删除临时形状。')
    sid=c.slide('page',1)
    b=c.edit('background','addShape',sid,kind='rectangle',left=30,top=30,width=650,height=430);bg=ref(b,'shapes',0,'id')
    c.style('background_style','formatShape',sid,bg,format={'fillColor':15921906,'lineVisible':False})
    text=c.text('label',sid,'星河知识库\n四周试点方案',index=1,left=80,top=80,width=450,height=120)
    s=c.edit('highlight','addShape',sid,kind='rectangle',left=100,top=250,width=180,height=70);highlight=ref(s,'shapes',2,'id')
    c.style('highlight_style','formatShape',sid,highlight,format={'fillColor':65535,'fillTransparency':0.25,'lineColor':255,'lineWidth':2})
    s=c.edit('resize','setShapeGeometry',sid,shapeId=highlight,geometry={'left':350,'top':250,'width':240,'height':90});c.shape(s,highlight,{'left':350,'top':250,'width':240,'height':90})
    s=c.edit('back','setShapeOrder',sid,shapeId=bg,position='back');c.shape(s,bg,{'zOrder':1})
    s=c.edit('front','setShapeOrder',sid,shapeId=text,position='front');c.shape(s,text,{'zOrder':3})
    s=c.edit('temporary','addShape',sid,kind='ellipse',left=50,top=360,width=50,height=50);temp=ref(s,'shapes',3,'id')
    s=c.edit('delete_temp','deleteShape',sid,shapeId=temp);c.check(s,['shapes','*','id'],'not_contains',temp);c.check(s,['slide','shapeCount'],'equals',3)
    c.finish()

    c=new(6,'幻灯片复制、移动、删除与身份保持','从 A/B/C/D 四页开始，复制 B 并改名“B副本”，移动到首页，删除 C；最终顺序 B副本、A、B、D，原 B 仍保留。')
    ids=[]
    for i,name in enumerate('ABCD',1):
        sid=c.slide('page'+name,i);ids.append(sid);c.text('title'+name,sid,'章节 '+name)
    s=c.edit('duplicate','duplicateSlide',ids[1]);copy=ref(s,'slides',2,'id');c.check(s,['slides'],'length',5)
    read=c.add('copy_info','getSlideInfo',slideId=copy);shape=ref(read,'shapes',0,'id')
    s=c.add('copy_title','setShapeText',slideId=copy,shapeId=shape,expectedToken=ref(read,'token'),text='章节 B副本');c.shape(s,shape,{'text':'章节 B副本'})
    ls=c.add('move_token','listSlides');s=c.add('move','moveSlide',slideId=copy,position=1,expectedToken=ref(ls,'token'));c.check(s,['slides','*','id'],'equals',[copy,*ids])
    s=c.edit('delete_C','deleteSlide',ids[2]);c.check(s,['slides','*','id'],'equals',[copy,ids[0],ids[1],ids[3]])
    c.finish()

    c=new(7,'三页图文表混排','三页均包含标题、PNG 图片、2×3 原生表格和备注；图片与表格位于独立区域，文字完整。每页内容不同且不串页。')
    for i in range(1,4):
        sid=c.slide('page%d'%i,i);title=c.text('title%d'%i,sid,'第%d部门进展'%i);c.font('font%d'%i,sid,title,size=26,bold=True)
        s=c.edit('image%d'%i,'addImage',sid,path=root+'/assets/sample.png',left=40,top=130,width=180,height=180)
        image_id=ref(s,'shapes',1,'id');c.shape(s,image_id,{'left':40,'top':130,'width':180,'height':180})
        values=[['部门','任务','状态'],['部门%d'%i,'条目%d'%i,'进行中']]
        s=c.edit('table%d'%i,'addTable',sid,values=values,left=260,top=140,width=400,height=150);c.check(s,['values'],'equals',values)
        c.notes_for('notes%d'%i,sid,'本页图片为测试素材，表格展示第%d部门数据。'%i)
    c.finish(pdf=True)

    c=new(8,'双表格与连续矩阵更新','一页放 8×5 明细表，另一页放 6×4 汇总表；分别更新两次，同样的行列结构保持，末列显示最终状态，不修改表格尺寸。')
    for index,(rows,cols) in enumerate(((8,5),(6,4)),1):
        sid=c.slide('page%d'%index,index);c.text('title%d'%index,sid,'表格验证 %d'%index)
        values=[['R%dC%d'%(r+1,k+1) for k in range(cols)] for r in range(rows)]
        s=c.edit('table%d'%index,'addTable',sid,values=values,left=35,top=130,width=650,height=280);tid=ref(s,'shapeId');c.check(s,['values'],'equals',values)
        for turn in range(1,3):
            r=c.add('read%d_%d'%(index,turn),'readTable',slideId=sid,shapeId=tid)
            values=[row[:-1]+['最终' if turn==2 else '处理中'] for row in values]
            s=c.add('update%d_%d'%(index,turn),'writeTable',slideId=sid,shapeId=tid,expectedToken=ref(r,'token'),values=values)
            c.check(s,['values'],'equals',values);c.check(s,['rows'],'equals',rows);c.check(s,['columns'],'equals',cols)
    c.finish()

    c=new(9,'中文查找替换与讲者备注','三段文本各有两处“待确认”，一共找到六处；逐个替换为“已确认”，保留 emoji 和换行。备注包含中文、英文及两行行动要求。')
    sid=c.slide('page',1);shapes=[]
    for i in range(3):shapes.append(c.text('box%d'%i,sid,'😀 部门%d：待确认\n负责人：待确认'%(i+1),index=i,top=25+135*i,height=120))
    s=c.add('find','findText',slideId=sid,text='待确认');c.check(s,['matches'],'length',6);c.check(s,['matches','*','length'],'all_equals',3)
    for i,shape in enumerate(shapes):
        s=c.edit('replace%d'%i,'replaceText',sid,shapeId=shape,find='待确认',replacement='已确认')
        c.check(s,['replacements'],'equals',2);c.shape(s,shape,{'text':'😀 部门%d：已确认\n负责人：已确认'%(i+1)})
    c.notes_for('notes',sid,'跟进负责人：研发 / Operations\n下周完成 FAQ 评审。')
    c.finish()

    seed=Case('P10-setup','ppt','准备 P10 已有文件','生成已保存的三页独立文稿副本，第二页标题含“草稿”。',root,setup=True)
    for i,title in enumerate(('项目概览','方案草稿','后续计划'),1):
        sid=seed.slide('page%d'%i,i);seed.text('title%d'%i,sid,title);seed.notes_for('notes%d'%i,sid,'原始备注%d'%i)
    seed.finish();setup.append(seed)
    c=new(10,'已有三页文稿的局部修订','只修订第二页为“方案定稿”，标题 30pt 加粗，更新备注和白色背景；保留第一、三页，原路径保存并导出 PDF。',seed.path('.pptx'))
    ls=c.add('slides','listSlides');c.check(ls,['slides'],'length',3);sid=ref(ls,'slides',1,'id')
    r=c.add('target','getSlideInfo',slideId=sid);c.check(r,['shapes',0,'text'],'equals','方案草稿');title=ref(r,'shapes',0,'id')
    s=c.add('rewrite','replaceText',slideId=sid,shapeId=title,expectedToken=ref(r,'token'),find='草稿',replacement='定稿');c.check(s,['replacements'],'equals',1);c.shape(s,title,{'text':'方案定稿'})
    c.font('font',sid,title,size=30,bold=True)
    c.notes_for('notes',sid,'正式评审已通过，按四周计划推进。');c.settings('settings',sid,backgroundColor=16777215,hidden=False)
    c.finish(pdf=True,save=True)

    c=new(11,'五页演示与 PPTX/PDF/PNG 交付','五页具备标题、正文、统一字体和备注；前三页分别导出 1280×720 PNG，并保存 PPTX、导出整份 PDF。导出不改变文稿状态。')
    ids=[]
    for i in range(1,6):
        sid=c.slide('page%d'%i,i);ids.append(sid)
        title=c.text('title%d'%i,sid,'星河知识库 · %d'%i)
        body=c.text('body%d'%i,sid,'目标：统一常见问题答案\n参与：研发、运营、客服\n阶段：第%d周检查点'%i,index=1,top=130,height=230)
        c.font('title_font%d'%i,sid,title,size=28,bold=True,latinName='Arial',eastAsianName='宋体')
        c.font('body_font%d'%i,sid,body,size=18)
        c.notes_for('notes%d'%i,sid,'讲解第%d页及其行动项。'%i)
    for i,sid in enumerate(ids[:3],1):c.png('page_image%d'%i,sid)
    c.finish(pdf=True)

    c=new(12,'复杂文稿中途失败后的停止','先完成三页标题和正文，再用过期 slide token 改写第一页。第二次改写被拒绝；前一次改写保留，后续备注、PNG、保存和 PDF 均未执行。')
    ids=[];heads=[]
    for i in range(1,4):
        sid=c.slide('page%d'%i,i);ids.append(sid)
        heads.append(c.text('title%d'%i,sid,'阶段%d'%i))
        c.text('body%d'%i,sid,'已完成的内容\n保留这些修改',index=1,top=120,height=200)
    r=c.add('old_token','getSlideInfo',slideId=ids[0])
    s=c.add('first_edit','setShapeText',slideId=ids[0],shapeId=heads[0],expectedToken=ref(r,'token'),text='已修改一次');c.shape(s,heads[0],{'text':'已修改一次'})
    c.add('stale_edit','setShapeText',slideId=ids[0],shapeId=heads[0],expectedToken=ref(r,'token'),text='不应出现')
    c.notes_for('must_not_notes',ids[0],'不应写入');c.png('must_not_export',ids[0]);c.finish(pdf=True);c.expect_failure('stale_edit','STALE_CONTENT')
    return cases,setup


def write_plan(root, output):
    windows=PureWindowsPath(root)
    if not windows.is_absolute() or '{' in root or '}' in root:raise ValueError('Use a concrete absolute Windows root, without placeholders')
    root=windows.as_posix().rstrip('/')
    all_cases=[];all_setup=[]
    for builder in (excel_cases,ppt_cases):
        cases,setup=builder(root);all_cases+=cases;all_setup+=setup
    # Complete validation happens before writing any plan artifact.
    expectations={c.id:c.expectation() for c in all_setup+all_cases}
    if output.exists():raise FileExistsError('Use a new plan directory: '+str(output))
    output.mkdir(parents=True)
    for name in ('requests','expected','assets'):(output/name).mkdir()
    make_png(output/'assets/sample.png')
    manifest={'status':'prepared_not_executed','windowsRoot':root,'actionTimeoutSeconds':60,
              'primaryCaseCount':len(all_cases),'setupTaskCount':len(all_setup),
              'executionPolicy':{'noAutomaticRetry':True,'stopOnUnexpectedSetup':True,'oneTaskAtATime':True,
                                 'directoriesPreparedByHarnessBeforeSubmission':True,'noRuntimeDirectoryRepair':True},
              'setup':[],'cases':[]}
    for c in all_setup+all_cases:
        (output/'requests'/(c.id+'.json')).write_text(json.dumps(c.request,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        expected=expectations[c.id]
        (output/'expected'/(c.id+'.json')).write_text(json.dumps(expected,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        manifest['setup' if c.setup else 'cases'].append({
            'id':c.id,'app':c.app,'title':c.title,'kind':expected['kind'],
            'request':'requests/'+c.id+'.json','expected':'expected/'+c.id+'.json',
            'plannedActionCount':len(expected['actions']),
            'assertionCount':sum(len(s['dataAssertions']) for s in expected['actions']),
            'dependsOn':['X10-setup'] if c.id=='X10' else ['P10-setup'] if c.id=='P10' else [],
            'humanExpected':c.effect,'initialConditions':c.initial})
    (output/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    lines=['# Excel / PPT 复杂执行测试：请求与预期','', '**状态：已准备，未执行。**', '',
           'Windows 运行根目录：`'+root+'`。更换轮次时重新生成整个目录，不能手动只改部分路径。', '',
           '| 用例 | 场景 | Action 数 | 预期 | 请求 | 逐步预期 |', '| --- | --- | ---: | --- | --- | --- |']
    for row in manifest['cases']:
        lines.append('| {id} | {title} | {plannedActionCount} | {kind} | [JSON]({request}) | [预期]({expected}) |'.format(**row))
    lines+=['','## 准备请求','']
    for row in manifest['setup']:lines.append('- '+row['id']+'：[JSON]('+row['request']+') / [预期]('+row['expected']+')。仅作为各自正式用例的独立输入，不计入 24 例分母。')
    lines+=['','## 人工粗审效果','']
    for row in manifest['cases']:lines+=['### '+row['id']+' '+row['title'],'',row['humanExpected'],'']
    (output/'README.md').write_text('\n'.join(lines)+'\n',encoding='utf-8')
    return manifest


def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--windows-root',required=True)
    parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args(argv)
    result=write_plan(args.windows_root,args.output)
    print(json.dumps({'status':result['status'],'cases':len(result['cases']),'setup':len(result['setup']),
                      'plannedActions':sum(c['plannedActionCount'] for c in result['cases'])},ensure_ascii=False))


if __name__=='__main__':main()
