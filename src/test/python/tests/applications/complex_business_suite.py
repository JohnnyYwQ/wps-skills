"""Sixty deterministic business Tasks; fixtures are separate from measured Tasks."""
import copy
from tests.applications.complex_plan import Case, excel_cases, ppt_cases
from tests.applications.action_suite import add_read_assertions
from tests.applications.native_acceptance import ref

INSPECT={'scope':{'kind':'document'},'limits':{'maxTextCharacters':16384,'maxParagraphs':256,'maxRuns':1024}}
BODY={'eastAsiaFontFamily':'宋体','westernFontFamily':'Arial','fontSizePt':12,'bold':False}
PARA={'alignment':'justify','lineSpacing':{'kind':'exact','points':20},'spaceAfterPt':6}

def block(text,heading=False,styled=False):
    b={'kind':'heading' if heading else 'paragraph','runs':[{'text':text,'format':dict(BODY,**({'fontSizePt':18,'bold':True} if heading else {}))}]}
    if heading:b['level']=1;b['format']={'alignment':'center'}
    elif styled:b['format']=copy.deepcopy(PARA)
    return b

def word_write(c,sid,texts,heading=False,styled=False,anchor=None,bold_last=False):
    bs=[block(t,heading and i==0,styled) for i,t in enumerate(texts)]
    if bold_last:bs[-1]['runs'][0]['format']['bold']=True
    c.add(sid,'writeContent',anchor=anchor or {'kind':'documentEnd'},blocks=bs)
    c.check(sid,['range','end'],'greater_than',0)
    # A range-scoped read has fixed text and format oracles, independent of returned success.
    read=sid+'_read'
    c.add(read,'inspectDocument',scope={'kind':'range','range':ref(sid,'range')},limits=INSPECT['limits'])
    c.check(read,['truncated'],'equals',False)
    c.check(read,['text'],'one_of',['\n'.join(texts),'\n'+'\n'.join(texts)])
    for i,b in enumerate(bs):
        # Paragraph observations include the paragraph mark. Character patches
        # apply only to run text, so inspect that exact text range for fonts.
        locate=sid+'_text%d'%i
        c.add(locate,'findContent',query={'scope':{'kind':'range','range':ref(sid,'range')},'text':texts[i],'caseSensitive':True,'wholeWord':False},limit=100)
        c.check(locate,['matches'],'length',sum(t.count(texts[i]) for t in texts))
        match_index=sum(t.count(texts[i]) for t in texts[:i])
        observed=locate+'_read'
        c.add(observed,'inspectDocument',scope={'kind':'range','range':ref(locate,'matches',match_index,'range')},limits=INSPECT['limits'])
        c.check(observed,['text'],'equals',texts[i]);c.check(observed,['truncated'],'equals',False)
        for k,v in b['runs'][0]['format'].items():c.check(observed,['paragraphs',0,'runs','*','format',k],'all_equals',v)
        if heading and i==0:c.check(observed,['paragraphs',0,'kind'],'equals','heading');c.check(observed,['paragraphs',0,'level'],'equals',1)
        for k,v in b.get('format',{}).items():c.check(observed,['paragraphs',0,'format',k],'equals',v)
    return read

def word_inspect(c,sid,expected=None):
    c.add(sid,'inspectDocument',**INSPECT);c.check(sid,['truncated'],'equals',False)
    if expected is not None:c.check(sid,['text'],'equals',expected)
    return sid

def word_replace(c,sid,old,new,count=1):
    q={'scope':{'kind':'document'},'text':old,'caseSensitive':True,'wholeWord':False}
    c.add(sid+'_find','findContent',query=q,limit=100);c.check(sid+'_find',['matches'],'length',count)
    replacement={'kind':'delete'} if new is None else {'kind':'blocks','blocks':[{'kind':'text','runs':[{'text':new}]}]}
    if count>1 and new is not None:replacement={'kind':'text','runs':[{'text':new}]}
    target={'kind':'range','range':ref(sid+'_find','matches',0,'range')} if count==1 else {'kind':'query','query':q,'expectedMatchCount':count}
    c.add(sid,'replaceContent',target=target,replacement=replacement);c.check(sid,['matchedCount'],'equals',count)

def word_table(c,sid,data):
    c.add(sid,'insertTable',anchor={'kind':'documentEnd'},data=data,headerRow=True);c.check(sid,['table','data'],'equals',data)

def word_image(c,sid):
    c.add(sid,'insertImage',anchor={'kind':'documentEnd'},source={'kind':'file','path':c.root+'/assets/sample.png'},placement={'kind':'inline'},size={'kind':'width','width':{'value':72,'unit':'pt'}},alternativeText={'kind':'decorative'})
    c.check(sid,['image','embedded'],'equals',True);c.check(sid,['image','size','width','value'],'close',72,0.1)

def word_layout(c,sid,orientation,index=None):
    r=word_inspect(c,sid+'_read')
    sections={'kind':'all','revision':ref(r,'revision')} if index is None else {'kind':'indexes','indexes':[index],'revision':ref(r,'revision')}
    c.add(sid,'setPageLayout',sections=sections,layout={'orientation':orientation,'margins':{k:{'value':2,'unit':'cm'} for k in ('top','bottom','left','right')}})
    c.check(sid,['sections','*','layout','orientation'],'all_equals',orientation)
    for k in ('top','bottom','left','right'):c.check(sid,['sections','*','layout','margins',k,'value'],'all_close',2*72/2.54,0.1)

def word_header(c):
    r=word_inspect(c,'header_read')
    c.add('header','setHeaderFooter',sections={'kind':'all','revision':ref(r,'revision')},updates=[{'area':a,'variant':'primary','operation':{'kind':'replace','text':t}} for a,t in [('header','星河项目资料'),('footer','内部正式版本')]])
    c.check('header',['stories','*','text'],'equals',['星河项目资料','内部正式版本'])

def word_cases(root):
    cases=[];seeds=[]
    def new(n,title,initial=None):
        existing=None
        if initial:
            seed=Case('CW%02d-setup'%n,'word','准备'+title,'固定已有文档',root,setup=True)
            word_write(seed,'seed',initial,bold_last=n==14);seed.finish();seeds.append(seed);existing=seed.path('.docx')
        c=Case('CW%02d'%n,'word',title,title,root,existing);cases.append(c);return c
    c=new(1,'项目启动通知');word_write(c,'notice',['项目启动通知','星河知识库于2026年10月12日启动。','研发、运营、客服共同参与。','首期四周，统一常见问题答案。'],True);word_replace(c,'date','2026年10月12日','2026年10月15日');word_inspect(c,'final','项目启动通知\n星河知识库于2026年10月15日启动。\n研发、运营、客服共同参与。\n首期四周，统一常见问题答案。');c.finish()
    c=new(2,'五章节项目报告')
    for i,title in enumerate(['试点概况','进展与收益','问题与风险','下一阶段计划','结论']):word_write(c,'chapter%d'%i,[title,['试点持续四周。','整理120条，确认102条。','18条待业务确认。','建立每周更新机制。','继续推进审核工作。'][i]],True,True)
    word_write(c,'append',['结论补充：由运营跟进责任落实。']);c.finish(pdf=True)
    c=new(3,'六主题长篇培训材料')
    for i,title in enumerate(['检索','辨别版本','提交问题','反馈错误','权限与保密','日常维护']):
        paragraphs=[title]+['%s练习%d：请根据岗位选择准确关键词，核对资料的适用范围和更新时间。遇到不同版本时记录来源，联系负责人确认后使用；不要传播未经确认的内容。完成练习后保存处理记录，供下一位同事复核。示例：使用业务名称与问题类别组合检索，先检查资料标题，再确认生效日期及负责人。练习完成后记录选用的版本、关键结论与仍待确认的问题，交由培训负责人统一答疑。'%(title,j) for j in range(1,6)]
        word_write(c,'chapter%d'%i,paragraphs,True)
    word_write(c,'conclusion',['培训结论：待发布。']);word_replace(c,'revise','待发布','已完成审核');c.finish()
    c=new(4,'双语维护通知');word_write(c,'bilingual',['系统维护通知 / Maintenance Notice','系统将于2026年10月15日22:00至23:00维护。','Please save your work before maintenance.'],True);word_replace(c,'revise','Please save your work before maintenance.','Please save your work in advance.');word_inspect(c,'final','系统维护通知 / Maintenance Notice\n系统将于2026年10月15日22:00至23:00维护。\nPlease save your work in advance.');c.finish()
    c=new(5,'会议纪要');word_write(c,'minutes',['知识库周例会纪要','本周完成18条答案审核。','仍有5条需要财务确认。','下周一由运营组织复核。'],True,True);word_write(c,'actions',['行动项：周五前汇总复核意见。'],styled=True);c.finish()
    c=new(6,'制度发布',['本文件为试行版。','试行版适用于首期培训。','反馈将在试行版期间收集。']);word_replace(c,'release','试行版','正式版',3);word_inspect(c,'final','本文件为正式版。\n正式版适用于首期培训。\n反馈将在正式版期间收集。');c.finish(save=True)
    c=new(7,'培训资料清稿',['培训资料草稿','[临时]本次培训面向新员工。','资料[临时]将在周五更新。']);word_replace(c,'delete','[临时]',None,2);word_replace(c,'title','培训资料草稿','培训资料正式版');word_inspect(c,'final','培训资料正式版\n本次培训面向新员工。\n资料将在周五更新。');c.finish(save=True)
    c=new(8,'阶段结论修订',['阶段报告','本阶段结论：仍需继续整理。','下一步安排：由运营组织复核。'])
    r=word_inspect(c,'read');c.add('replace','replaceContent',target={'kind':'range','range':ref(r,'paragraphs',1,'range')},replacement={'kind':'blocks','blocks':[block('本阶段结论：试点已完成，建立每周更新机制。',styled=True)]});c.check('replace',['matchedCount'],'equals',1)
    r=word_inspect(c,'final','阶段报告\n本阶段结论：试点已完成，建立每周更新机制。\n下一步安排：由运营组织复核。')
    for k,v in PARA.items():c.check(r,['paragraphs',1,'format',k],'equals',v)
    c.finish(save=True)
    c=new(9,'报告摘要更新',['试点报告','四周整理120条，确认102条，剩余18条待确认。','总结','待补充总结。']);word_inspect(c,'read');word_replace(c,'summary','待补充总结。','四周试点已整理120条答案，102条完成确认，下一阶段优先处理18条待确认内容。');word_inspect(c,'final','试点报告\n四周整理120条，确认102条，剩余18条待确认。\n总结\n四周试点已整理120条答案，102条完成确认，下一阶段优先处理18条待确认内容。');c.finish(save=True)
    c=new(10,'工作安排');word_write(c,'intro',['第一阶段工作安排','以下为第一阶段的责任分工。'],True);word_table(c,'table',[['任务','负责人','截止日期'],['资料整理','张敏','2026-10-12'],['答案审核','李明','2026-10-16'],['发布培训','王芳','2026-10-20']]);word_write(c,'tail',['各负责人每周反馈进度。']);c.finish()
    c=new(11,'双节项目说明');word_write(c,'first',['项目说明','本部分采用纵向页面。'],True);c.add('break','insertBreak',anchor={'kind':'documentEnd'},type='sectionNextPage');c.check('break',['break','sectionCountAfter'],'equals',2);word_write(c,'second',['数据展示','本部分采用横向页面。'],True);word_layout(c,'portrait','portrait',0);word_layout(c,'landscape','landscape',1);c.finish()
    c=new(12,'成果简报');word_write(c,'intro',['试点成果简报','整理120条，确认102条，待确认18条。'],True);word_table(c,'table',[['项目','数量','说明'],['总数','120','已整理'],['确认','102','可发布'],['待确认','18','优先审核']]);c.finish(pdf=True)
    c=new(13,'图文操作指引');word_write(c,'intro',['知识库操作指引','步骤一：搜索问题。'],True);word_image(c,'image1');word_write(c,'second',['步骤二：核对版本。']);word_image(c,'image2');word_write(c,'third',['步骤三：记录反馈。']);c.finish()
    c=new(14,'已有资料续写',['培训资料','重要提醒：先阅读目录。']);word_inspect(c,'read');word_write(c,'append',['新增培训安排','周五参加演练，完成后提交反馈。'],True);c.finish(save=True)
    c=new(15,'风险章节扩充',['项目报告','本周完成资料整理。','风险与应对','部分答案仍待确认。'])
    q={'scope':{'kind':'document'},'text':'部分答案仍待确认。','caseSensitive':True,'wholeWord':False};c.add('find','findContent',query=q,limit=10);c.check('find',['matches'],'length',1)
    word_write(c,'risks',['风险一：口径不一致；对策：统一审核。','风险二：资料过期；对策：每周更新。','风险三：权限错误；对策：发布前复核。'],anchor={'kind':'before','range':ref('find','matches',0,'range')});c.finish(save=True)
    c=new(16,'统一页眉页脚',['培训正文第一段。','培训正文第二段。']);word_write(c,'append',['补充：按正式流程反馈问题。']);word_header(c);c.finish(save=True)
    c=new(17,'打印版排版',['打印版资料','请按章节核对信息。']);word_layout(c,'layout','landscape');word_write(c,'append',['打印前复核日期与版本。']);c.finish(save=True)
    c=new(18,'归档副本',['归档材料','版本说明：草稿']);word_replace(c,'version','草稿','正式归档版');word_inspect(c,'final','归档材料\n版本说明：正式归档版');c.finish(pdf=True)
    c=new(19,'合同多点修订',['项目名：星河试点','日期：2026-10-12','金额：10000元'])
    for sid,old,newtext in [('name','星河试点','星河正式项目'),('date','2026-10-12','2026-10-20'),('amount','10000元','12000元')]:word_replace(c,sid,old,newtext)
    word_write(c,'note',['附注：以双方确认版本为准。']);word_inspect(c,'final','项目名：星河正式项目\n日期：2026-10-20\n金额：12000元\n附注：以双方确认版本为准。');c.finish(save=True)
    c=new(20,'综合发布材料');word_write(c,'intro',['发布计划','发布日期：待定','[内部待审]本次发布覆盖三个部门。'],True);word_replace(c,'date','待定','2026-10-20');word_replace(c,'marker','[内部待审]',None);word_table(c,'table',[['部门','任务'],['研发','技术复核'],['运营','发布'],['客服','培训']]);word_image(c,'image');word_header(c);c.finish(pdf=True)
    return cases,seeds


def additional_excel(root):
    cases=[];seeds=[]
    def new(n,title,existing=None):
        c=Case('CE%02d'%n,'excel',title,title,root,existing);cases.append(c);return c
    c=new(12,'季度销售汇总');sh=c.sheet('一月');totals=[]
    for m,name in enumerate(['一月','二月','三月'],1):
        if m>1:c.structure('month%d'%m,'addWorksheet',sh,name=name)
        c.region('data%d'%m,'writeRange',name,'A1:B13',values=[['订单','金额']]+[['订单%02d'%j,j*m*10] for j in range(1,13)])
        c.region('sum%d'%m,'setFormulas',name,'B15',formulas=[['=SUM(B2:B13)']]);s=c.region('calc%d'%m,'calculateRange',name,'B15');c.values(s,[[780*m]]);totals.append(ref(s,'cells',0,0,'value'))
    c.structure('quarter','addWorksheet',sh,name='季度');c.region('quarter_data','writeRange','季度','A1:B4',values=[['月份','金额'],*[[name,v] for name,v in zip(['一月','二月','三月'],totals)]]);c.region('sum','setFormulas','季度','B6',formulas=[['=SUM(B2:B4)']]);s=c.region('calc','calculateRange','季度','B6');c.values(s,[[4680]]);c.finish()
    c=new(13,'库存盘点');sh=c.sheet('库存');data=[['物料','期初','入库','出库','结存']]+[['M%02d'%i,100,i*3,i,None] for i in range(1,11)];c.region('data','writeRange',sh,'A1:E11',values=data);c.region('formula','setFormulas',sh,'E2:E11',formulas=[['=B%d+C%d-D%d'%(r,r,r)] for r in range(2,12)]);s=c.region('calc','calculateRange',sh,'E2:E11');c.values(s,[[100+2*i] for i in range(1,11)]);s=c.region('sort','sortRange',sh,'A1:E11',column=5,order='descending',header=True);c.values(s,[data[0]]+[[*data[i][:-1],100+2*i] for i in range(10,0,-1)]);c.structure('snapshot','addWorksheet',sh,name='盘点');c.copy('copy',sh,'A1:E11','盘点','A1:E11');c.finish()
    c=new(14,'费用报销');sh=c.sheet('报销');data=[['单号','数量','单价','金额','审批']]+[['F%02d'%i,i,10,None,'已审批' if i%2 else '待审批'] for i in range(1,16)];c.region('data','writeRange',sh,'A1:E16',values=data);c.region('formulas','setFormulas',sh,'D2:D16',formulas=[['=B%d*C%d'%(r,r)] for r in range(2,17)]);s=c.region('calculate','calculateRange',sh,'D2:D16');c.values(s,[[i*10] for i in range(1,16)]);c.region('totals','setFormulas',sh,'G2:G3',formulas=[['=SUM(D2:D16)'],['=SUMIF(E2:E16,"已审批",D2:D16)']]);s=c.region('calc','calculateRange',sh,'G2:G3');c.values(s,[[1200],[640]]);s=c.region('filter','filterRange',sh,'A1:E16',column=5,value='已审批');c.check(s,['cells','*',0,'rowHidden'],'equals',[False]+[i%2==0 for i in range(1,16)]);c.structure('publish','addWorksheet',sh,name='发布');c.region('write_total','writeRange','发布','A1:B2',values=[['总额','已审批'],[ref('calc','cells',0,0,'value'),ref('calc','cells',1,0,'value')]]);c.finish()
    c=new(15,'培训成绩');sh=c.sheet('成绩');data=[['姓名','测验一','测验二','平均','结论']]+[['学员%02d'%i,40+2*i,50+2*i,None,None] for i in range(1,21)];c.region('data','writeRange',sh,'A1:E21',values=data);c.region('formulas','setFormulas',sh,'D2:E21',formulas=[['=AVERAGE(B%d:C%d)'%(r,r),'=IF(D%d>=60,"合格","待补训")'%r] for r in range(2,22)]);s=c.region('calc','calculateRange',sh,'D2:E21');c.values(s,[[45+2*i,'合格' if i>=8 else '待补训'] for i in range(1,21)]);s=c.region('rank','sortRange',sh,'A1:E21',column=4,order='descending',header=True);c.values(s,[data[0]]+[[*data[i][:3],45+2*i,'合格' if i>=8 else '待补训'] for i in range(20,0,-1)]);c.region('count','setFormulas',sh,'G2',formulas=[['=COUNTIF(E2:E21,"合格")']]);s=c.region('count_calc','calculateRange',sh,'G2');c.values(s,[[13]]);c.finish()
    c=new(16,'项目预算');sh=c.sheet('预算');data=[['项目','预算','实际','差额','执行率','状态']]+[['项目%d'%i,1000,800+i*100,None,None,None] for i in range(1,7)];c.region('data','writeRange',sh,'A1:F7',values=data);c.region('formula','setFormulas',sh,'D2:F7',formulas=[['=B%d-C%d'%(r,r),'=C%d/B%d'%(r,r),'=IF(D%d<0,"超支","正常")'%r] for r in range(2,8)]);s=c.region('calc','calculateRange',sh,'D2:F7');c.values(s,[[200-i*100,(800+i*100)/1000,'超支' if i>2 else '正常'] for i in range(1,7)]);c.region('money','formatRange',sh,'B2:D7',format={'numberFormat':'0.00'});c.region('ratio','formatRange',sh,'E2:E7',format={'numberFormat':'0.0%'});s=c.region('width','autoFitColumns',sh,'A1:F7');c.check(s,['cells',0,0,'columnWidth'],'greater_than',0);c.region('sum','setFormulas',sh,'B9:D9',formulas=[['=SUM(B2:B7)','=SUM(C2:C7)','=SUM(D2:D7)']]);s=c.region('sum_calc','calculateRange',sh,'B9:D9');c.values(s,[[6000,6900,-900]]);c.finish()
    c=new(17,'通讯录清理');sh=c.sheet('通讯录');rows=[['编号','姓名','部门']]+[['%03d'%i,'同事%d'%i,'运营组' if i%2 else '研发'] for i in range(1,11)];c.region('data','writeRange',sh,'A1:C11',values=rows)
    c.region('replace','replaceInRange',sh,'A1:C11',text='运营组',replacement='运营',matchCase=True,wholeCell=True)
    updated=[rows[0]]+[[r[0],r[1],r[2].replace('运营组','运营')] for r in rows[1:]]
    s=c.region('sort','sortRange',sh,'A1:C11',column=1,order='descending',header=True);c.values(s,[updated[0]]+updated[:0:-1]);c.finish()
    seed=Case('CE18-setup','excel','准备周报','固定上周数据',root,setup=True);sh=seed.sheet('上周');seed.region('data','writeRange',sh,'A1:B3',values=[['指标','数量'],['整理',20],['审核',18]]);seed.finish();seeds.append(seed)
    c=new(18,'周报滚动归档',seed.path('.xlsx'));c.structure('copy','copyWorksheet','上周',name='本周');c.region('update','writeRange','本周','B2:B3',values=[[30],[25]]);c.structure('move','moveWorksheet','本周',index=1);s=c.add('original','readRange',sheet='上周',address='A1:B3');c.values(s,[['指标','数量'],['整理',20],['审核',18]]);s=c.add('final_order','listWorksheets',offset=0,limit=100);c.check(s,['worksheets','*','name'],'equals',['本周','上周']);c.finish(save=True)
    c=new(19,'对账交付');sh=c.sheet('对账');rows=[['订单','应收','实收','差额','状态']]+[['O%d'%i,i*100,i*100-(10 if i%3==0 else 0),None,None] for i in range(1,13)];c.region('data','writeRange',sh,'A1:E13',values=rows);c.region('formula','setFormulas',sh,'D2:E13',formulas=[['=B%d-C%d'%(r,r),'=IF(D%d=0,"一致","异常")'%r] for r in range(2,14)]);s=c.region('calc','calculateRange',sh,'D2:E13');c.values(s,[[10 if i%3==0 else 0,'异常' if i%3==0 else '一致'] for i in range(1,13)]);s=c.region('filter','filterRange',sh,'A1:E13',column=5,value='异常');c.check(s,['cells','*',0,'rowHidden'],'equals',[False]+[i%3!=0 for i in range(1,13)]);c.structure('snapshot','addWorksheet',sh,name='快照');c.copy('copy',sh,'A1:E13','快照','A1:E13');c.finish()
    c=new(20,'综合经营简报');sh=c.sheet('明细');rows=[['单号','部门','收入']]+[['B%02d'%i,'研发' if i%2 else '运营',i*100] for i in range(1,13)];c.region('data','writeRange',sh,'A1:C13',values=rows);s=c.region('sort','sortRange',sh,'A1:C13',column=3,order='descending',header=True);c.values(s,[rows[0]]+rows[:0:-1]);c.region('filter','filterRange',sh,'A1:C13',column=2,value='研发');s=c.region('clear','clearFilter',sh,'A1:C13');c.check(s,['cells','*',0,'rowHidden'],'all_equals',False);c.region('formula','setFormulas',sh,'E2:E4',formulas=[['=SUMIF(B2:B13,"研发",C2:C13)'],['=SUMIF(B2:B13,"运营",C2:C13)'],['=SUM(C2:C13)']]);s=c.region('calc','calculateRange',sh,'E2:E4');c.values(s,[[3600],[4200],[7800]]);c.structure('report','addWorksheet',sh,name='简报');c.region('title','writeRange','简报','A1:C1',values=[['经营简报',None,None]]);c.region('merge','mergeRange','简报','A1:C1');c.region('summary','writeRange','简报','A3:B5',values=[['研发',ref('calc','cells',0,0,'value')],['运营',ref('calc','cells',1,0,'value')],['总计',ref('calc','cells',2,0,'value')]]);c.finish(pdf=True)
    return cases,seeds


def additional_ppt(root):
    cases=[];seeds=[]
    def new(n,title,existing=None):
        c=Case('CP%02d'%n,'ppt',title,title,root,existing);cases.append(c);return c
    def page(c,key,pos,title,body):
        sid=c.slide(key,pos);h=c.text(key+'_title',sid,title);c.font(key+'_font',sid,h,size=26,bold=True)
        c.text(key+'_body',sid,body,index=1,top=120,height=220)
        c.notes_for(key+'_notes',sid,'讲解：'+title)
        return sid
    c=new(12,'季度复盘')
    for i,title in enumerate(['季度概览','核心指标','问题复盘','改进计划','行动项'],1):
        sid=page(c,'page%d'%i,i,title,['季度目标已完成。','整理120条，确认102条。','18条需要复核。','每周跟踪审核进度。','运营负责安排下周复核。'][i-1])
        if i==2:
            values=[['指标','数量'],['整理','120'],['确认','102'],['待确认','18']];s=c.edit('metrics','addTable',sid,values=values,left=50,top=350,width=600,height=130);c.check(s,['values'],'equals',values)
    c.finish()
    c=new(13,'产品介绍')
    for i,title in enumerate(['产品定位','检索功能','审核功能','使用收益'],1):
        sid=page(c,'page%d'%i,i,title,'统一知识入口，支持日常业务查询。')
        s=c.edit('image%d'%i,'addImage',sid,path=root+'/assets/sample.png',left=480,top=280,width=160,height=120);shape=ref(s,'shapes',2,'id');c.shape(s,shape,{'left':480,'top':280,'width':160,'height':120})
    c.finish()
    c=new(14,'培训课件');ids=[]
    for i,title in enumerate(['培训目标','登录','检索','反馈','总结'],1):
        sid=page(c,'page%d'%i,i,title,'步骤%d：按说明完成操作并保留记录。'%i);ids.append(sid)
        shape=c.text('tip%d'%i,sid,'提示：核对当前版本',index=2,left=60,top=370,width=570,height=50);c.style('tip_style%d'%i,'formatShape',sid,shape,format={'fillColor':65535,'lineVisible':False})
    ls=c.add('order','listSlides');s=c.add('move','moveSlide',slideId=ids[3],position=3,expectedToken=ref(ls,'token'));c.check(s,['slides','*','id'],'equals',[ids[0],ids[1],ids[3],ids[2],ids[4]]);c.finish()
    c=new(15,'双栏方案比较')
    for i,title in enumerate(['实施范围','资源投入','验收方案'],1):
        sid=c.slide('page%d'%i,i);c.text('title%d'%i,sid,title)
        ids=[]
        for j,label in enumerate(['方案甲','方案乙']):
            shape=c.text('column%d_%d'%(i,j),sid,label+'：分阶段实施',index=j+1,left=35+350*j,top=120,width=300,height=80);ids.append(shape)
        s=c.edit('align%d'%i,'alignShapes',sid,shapeIds=ids,alignment='top')
        for shape in ids:c.shape(s,shape,{'top':120,'width':300,'height':80})
        for j,label in enumerate(['方案甲','方案乙']):
            values=[['方案','周期'],[label,str(i+j+2)+'周']];s=c.edit('table%d_%d'%(i,j),'addTable',sid,values=values,left=35+350*j,top=240,width=300,height=150);c.check(s,['values'],'equals',values)
    c.finish()
    c=new(16,'路线图');sid=c.slide('page',1);ids=[]
    for i,title in enumerate(['调研','开发','试点','发布']):
        shape=c.text('stage%d'%i,sid,title,index=i,left=30+i*175,top=130,width=140,height=70);ids.append(shape)
        c.style('stage_style%d'%i,'formatShape',sid,shape,format={'fillColor':[15921906,13434879,65535,13421823][i],'lineWidth':1})
    s=c.edit('distribute','distributeShapes',sid,shapeIds=ids,direction='horizontal')
    for i,shape in enumerate(ids):c.shape(s,shape,{'left':30+i*175,'top':130})
    c.text('description',sid,'四阶段依次验收，由研发和运营共同推进。',index=4,top=280,height=120);c.png('roadmap',sid);c.finish(pdf=True)
    c=new(17,'组织职责');sid=c.slide('page',1);ids=[]
    for i,(dept,duty) in enumerate([('研发','技术复核'),('运营','资料发布'),('客服','用户培训')]):
        shape=c.text('dept%d'%i,sid,dept+'：'+duty,index=i,left=40+i*230,top=140,width=200,height=100);ids.append(shape);c.style('style%d'%i,'formatShape',sid,shape,format={'fillColor':15921906,'lineWidth':1})
    s=c.edit('align','alignShapes',sid,shapeIds=ids,alignment='top')
    for shape in ids:c.shape(s,shape,{'top':140})
    s=c.edit('distribute','distributeShapes',sid,shapeIds=ids,direction='horizontal')
    for i,shape in enumerate(ids):c.shape(s,shape,{'left':40+230*i,'width':200,'height':100})
    s=c.edit('front','setShapeOrder',sid,shapeId=ids[1],position='front');c.shape(s,ids[1],{'zOrder':3});c.notes_for('notes',sid,'责任人按部门对应，不使用自动组织图。');c.finish()
    seed=Case('CP18-setup','ppt','准备客户提案','固定原始提案',root,setup=True)
    for i,title in enumerate(['客户甲概览','客户甲方案','内部讨论'],1):sid=seed.slide('page%d'%i,i);seed.text('title%d'%i,sid,title)
    seed.finish();seeds.append(seed)
    c=new(18,'客户提案派生',seed.path('.pptx'));ls=c.add('slides','listSlides');c.check(ls,['slides'],'length',3);ids=[ref(ls,'slides',i,'id') for i in range(3)]
    s=c.edit('duplicate','duplicateSlide',ids[1]);duplicate=ref(s,'slides',2,'id');c.check(s,['slides'],'length',4)
    for i,sid in enumerate([ids[0],ids[1],duplicate]):
        r=c.add('read%d'%i,'getSlideInfo',slideId=sid);shape=ref(r,'shapes',0,'id');s=c.add('replace%d'%i,'replaceText',slideId=sid,shapeId=shape,expectedToken=ref(r,'token'),find='客户甲',replacement='客户乙');c.check(s,['replacements'],'equals',1);c.shape(s,shape,{'text':'客户乙概览' if i==0 else '客户乙方案'})
    s=c.edit('delete_internal','deleteSlide',ids[2]);c.check(s,['slides','*','id'],'equals',[ids[0],ids[1],duplicate]);c.finish()
    c=new(19,'会务材料定稿');sid=page(c,'page',1,'会议议程','嘉宾：张敏、李明。')
    values=[['时间','议程'],['09:00','开场'],['09:30','待定']];s=c.edit('agenda','addTable',sid,values=values,left=40,top=270,width=600,height=160);table=ref(s,'shapeId');c.check(s,['values'],'equals',values)
    r=c.add('table_read','readTable',slideId=sid,shapeId=table);values=copy.deepcopy(values);values[-1][1]='专题讨论';s=c.add('table_update','writeTable',slideId=sid,shapeId=table,expectedToken=ref(r,'token'),values=values);c.check(s,['values'],'equals',values)
    placeholder=c.text('placeholder',sid,'临时占位',index=3,top=450,height=30);s=c.edit('delete_placeholder','deleteShape',sid,shapeId=placeholder);c.check(s,['shapes','*','id'],'not_contains',placeholder);c.check(s,['slide','shapeCount'],'equals',3);c.notes_for('final_notes',sid,'主持人提前十分钟到场。');c.finish()
    c=new(20,'综合成果汇报');ids=[]
    for i,title in enumerate(['成果概览','数据','案例','问题','计划','结论'],1):
        sid=page(c,'page%d'%i,i,title,'本页固定内容：'+title);ids.append(sid)
        if i==2:
            values=[['整理','确认','待确认'],['120','102','18']];s=c.edit('table','addTable',sid,values=values,left=50,top=350,width=600,height=120);c.check(s,['values'],'equals',values)
        if i==3:
            s=c.edit('image','addImage',sid,path=root+'/assets/sample.png',left=480,top=280,width=160,height=120);c.shape(s,ref(s,'shapes',2,'id'),{'width':160,'height':120})
    ls=c.add('order','listSlides');s=c.add('move','moveSlide',slideId=ids[4],position=4,expectedToken=ref(ls,'token'));c.check(s,['slides','*','id'],'equals',ids[:3]+[ids[4],ids[3],ids[5]])
    c.settings('highlight',ids[5],backgroundColor=15921906,hidden=False);c.png('overview',ids[0]);c.finish(pdf=True)
    return cases,seeds


def build_cases(app,root):
    if app=='word':cases,seeds=word_cases(root)
    else:
        old,seeds=(excel_cases if app=='excel' else ppt_cases)(root)
        cases=old[:11]
        # Rename identities; keep original generated fixture locators intact.
        for i,c in enumerate(cases,1):c.id=('CE' if app=='excel' else 'CP')+'%02d'%i
        more,extra=(additional_excel if app=='excel' else additional_ppt)(root);cases+=more;seeds+=extra
        if app=='ppt':
            c=cases[9]
            for index,title in [(0,'项目概览'),(2,'后续计划')]:
                s=c.add('preserved_page%d'%index,'getSlideInfo',slideId=ref('slides','slides',index,'id'))
                c.check(s,['shapes',0,'text'],'equals',title)
        # Cover application metadata and utility Actions in a relevant business case.
        c=cases[0];sid=c.add('application_info','getWorkbookInfo' if app=='excel' else 'getPresentationInfo')
        c.check(sid,['worksheetCount' if app=='excel' else 'slideCount'],'greater_than',0)
        if app=='excel':
            c.region('integer_format','formatRange','研发','C2:C13',format={'numberFormat':'0'})
        else:
            c=cases[12];r=c.add('rename_read','getSlideInfo',slideId=ref('page1','slides',0,'id'));shape=ref(r,'shapes',0,'id');s=c.add('rename_title','renameShape',slideId=ref(r,'slide','id'),shapeId=shape,expectedToken=ref(r,'token'),name='产品定位标题');c.shape(s,shape,{'name':'产品定位标题'})
    for c in seeds+cases:
        add_read_assertions(c)
        # Observation identity assertions for every read used to acquire references.
        for s in c.request['steps']:
            name=s['address']['action'];sid=s['id']
            if not c.checks.get(sid):
                paths={'listWorksheets':['worksheets'],'listSlides':['token'],'getSlideInfo':['token'],'getSlideNotes':['token'],'getSlideSettings':['token'],'readTable':['token'],'inspectDocument':['revision']}
                if name in paths:c.check(sid,paths[name],'not_null',None)
                else:raise ValueError('Missing explicit data assertion: '+c.id+'/'+sid+'/'+name)
        c.expectation()
    assert len(cases)==20
    return seeds+cases
