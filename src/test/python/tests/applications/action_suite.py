"""Prepare Action acceptance at the real Task Request/Response interface."""
import copy
from wps_skills.client.applications import contracts_for, compile_request
from wps_skills.client.task_client import _preflight
from tests.applications.complex_plan import Case, excel_cases, ppt_cases
from tests.applications.native_acceptance import ref

INSPECT = {'scope': {'kind': 'document'}, 'limits': {'maxTextCharacters': 4096, 'maxParagraphs': 128, 'maxRuns': 512}}


def paragraph(text):
    return {'kind': 'paragraph', 'runs': [{'text': text}]}


def word_cases(root):
    cases = []
    c = Case('W01', 'word', '文字、查找、范围替换和保存', '旧值替换为新值；回读与保存文件一致。', root)
    c.add('write', 'writeContent', anchor={'kind':'documentEnd'}, blocks=[paragraph('执行基线：旧值。')])
    c.check('write', ['range','start'], 'equals', 0)
    c.check('write', ['range','end'], 'greater_than', 0)
    c.add('find', 'findContent', query={'scope':{'kind':'document'}, 'text':'旧值','caseSensitive':True,'wholeWord':False}, limit=50)
    c.check('find',['matches'],'length',1)
    c.add('replace','replaceContent',target={'kind':'range','range':ref('find','matches',0,'range')},replacement={'kind':'blocks','blocks':[{'kind':'text','runs':[{'text':'新值'}]}]})
    c.check('replace',['matchedCount'],'equals',1)
    c.add('inspect','inspectDocument',**INSPECT)
    c.check('inspect',['text'],'equals','执行基线：新值。')
    c.check('inspect',['truncated'],'equals',False)
    c.finish(pdf=True);cases.append(c)
    reopened = Case('W02','word','重新打开并原路径保存','重新读取 W01 内容，追加一段后保存原路径。',root,c.path('.docx'))
    reopened.add('inspect','inspectDocument',**INSPECT)
    reopened.check('inspect',['text'],'equals','执行基线：新值。')
    reopened.add('append','writeContent',anchor={'kind':'documentEnd'},blocks=[paragraph('重新打开后追加。')])
    reopened.check('append',['range','end'],'greater_than',0)
    reopened.finish(save=True);cases.append(reopened)
    c=Case('W03','word','表格、图片、分页、页眉页脚和页面设置','2×2 表格、嵌入图片、分页、页眉页脚与横向页面。',root)
    c.add('write','writeContent',anchor={'kind':'documentEnd'},blocks=[paragraph('表格与图像测试')])
    values=[['名称','数量'],['测试', '2']]
    c.add('table','insertTable',anchor={'kind':'documentEnd'},data=values,headerRow=True)
    c.check('table',['table','data'],'equals',values)
    c.add('image','insertImage',anchor={'kind':'documentEnd'},source={'kind':'file','path':root+'/assets/sample.png'},placement={'kind':'inline'},size={'kind':'width','width':{'value':36,'unit':'pt'}},alternativeText={'kind':'decorative'})
    c.check('image',['image','embedded'],'equals',True)
    c.check('image',['image','size','width','value'],'close',36,0.1)
    c.add('break','insertBreak',anchor={'kind':'documentEnd'},type='page')
    c.check('break',['break','type'],'equals','page')
    c.check('break',['break','sectionCountBefore'],'equals',ref('break','break','sectionCountAfter'))
    c.add('inspect','inspectDocument',**INSPECT)
    c.add('layout','setPageLayout',sections={'kind':'all','revision':ref('inspect','revision')},layout={'orientation':'landscape'})
    c.check('layout',['sections','*','layout','orientation'],'all_equals','landscape')
    c.add('header','setHeaderFooter',sections={'kind':'all','revision':ref('layout','revisionAfter')},updates=[{'area':'header','variant':'primary','operation':{'kind':'replace','text':'验收页眉'}},{'area':'footer','variant':'primary','operation':{'kind':'replace','text':'验收页脚'}}])
    c.check('header',['stories','*','text'],'equals',['验收页眉','验收页脚'])
    c.finish(pdf=True);cases.append(c)
    return cases


def add_read_assertions(case):
    """Explicit request/result checks supplement scenario-specific literal oracles."""
    for step in case.request['steps']:
        name, params, sid = step['address']['action'], step['params'], step['id']
        if name == 'readRange':
            for key in ('sheet','address'):case.check(sid,[key],'equals',params[key])
        elif name == 'getWorksheetInfo':case.check(sid,['name'],'equals',params['sheet'])
        elif name == 'getShapeStyle':
            case.check(sid,['slideId'],'equals',params['slideId']);case.check(sid,['shape','id'],'equals',params['shapeId'])
        elif name in ('getSlideSettings','getSlideNotes'):case.check(sid,['slideId'],'equals',params['slideId'])
        elif name == 'readTable':
            case.check(sid,['slideId'],'equals',params['slideId']);case.check(sid,['shapeId'],'equals',params['shapeId'])


def build_cases(app, root):
    if app == 'word':
        cases = word_cases(root)
    else:
        cases, setup = (excel_cases if app=='excel' else ppt_cases)(root)
        # Dependencies precede their consumer; they are evidence-bearing setup Tasks.
        cases = [s for c in cases for s in ([x for x in setup if x.id[:3]==c.id]+[c])]
        extra = Case('X13' if app=='excel' else 'P13', app, '补齐正式 Action 覆盖', '验证应用信息及遗漏的命名/列宽操作。', root)
        if app=='excel':
            extra.add('info','getWorkbookInfo');extra.check('info',['worksheetCount'],'greater_than',0)
            sh=extra.sheet('列宽')
            extra.region('data','writeRange',sh,'A1:B2',values=[['列宽测试长标题','B'],['中文 mixed',2]])
            step=extra.region('fit','autoFitColumns',sh,'A1:B2')
            extra.check(step,['cells',0,0,'columnWidth'],'greater_than',0)
        else:
            extra.add('info','getPresentationInfo');extra.check('info',['slideCount'],'equals',0)
            sid=extra.slide('page',1);shape=extra.text('title',sid,'重命名测试')
            s=extra.edit('rename','renameShape',sid,shapeId=shape,name='验收标题')
            extra.shape(s,shape,{'name':'验收标题'})
        extra.finish()
        # Negative scene is last so its unsaved test document does not affect positives.
        cases.insert(next((i for i,c in enumerate(cases) if c.fail),len(cases)),extra)
    for case in cases:
        add_read_assertions(case)
        case.expectation()  # Compile and validate every positive/expected-execution-failure request.
    return cases


def coverage(app, cases):
    matrix={c.name:{'positiveCases':[], 'dataAssertionCases':[], 'invalidParameterCase':'invalid-'+c.name} for c in contracts_for(app).contracts}
    for case in cases:
        for step in case.expectation()['actions']:
            if step['state']!='succeeded':continue
            row=matrix[step['address']['action']]
            if case.id not in row['positiveCases']:row['positiveCases'].append(case.id)
            if step['dataAssertions'] and case.id not in row['dataAssertionCases']:row['dataAssertionCases'].append(case.id)
    missing=[name for name,row in matrix.items() if not row['positiveCases'] or not row['dataAssertionCases']]
    if missing:raise ValueError('Uncovered production Actions: '+', '.join(missing))
    return matrix


def invalid_cases(app,cases):
    """One known-invalid extra parameter per Action, rejected before document access."""
    requests={}
    for case in cases:
        for original in [case.request['document'],*case.request['steps'],*case.request['completion']]:
            name=original['address']['action']
            if name in requests:continue
            request=copy.deepcopy(case.request)
            steps=[request['document'],*request['steps'],*request['completion']]
            target=next(s for s in steps if s['address']['action']==name)
            target['params']['__acceptance_invalid_field__']=True
            try:_preflight(compile_request(request,app),contracts_for(app))
            except ValueError:pass
            else:raise ValueError('Negative fixture is not rejected: '+name)
            requests[name]=request
    return requests
