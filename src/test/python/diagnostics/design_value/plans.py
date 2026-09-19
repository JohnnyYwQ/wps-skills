"""Fixed Task inputs with literal, independent document oracles."""
from dataclasses import dataclass
from tests.applications.complex_plan import Case
from tests.applications.action_suite import paragraph

BASE = 'DESIGN_BASE_中文'
MARK = 'DESIGN_EFFECT_中文'
SHEET = '实验'
EXT = {'word': '.docx', 'excel': '.xlsx', 'ppt': '.pptx'}


@dataclass
class Plan:
    request: dict
    path: str


def seed(app, root, identifier):
    c = Case(identifier, app, '设计实验前置文档', BASE, root)
    if app == 'word':
        c.add('base', 'writeContent', anchor={'kind': 'documentEnd'}, blocks=[paragraph(BASE)])
    elif app == 'excel':
        c.sheet(SHEET)
        c.region('base', 'writeRange', SHEET, 'A1', values=[[BASE]])
    else:
        sid = c.slide('base', 1)
        c.text('base_text', sid, BASE)
    c.finish()
    return Plan(c.request, c.path(EXT[app]))


def edit(app, root, identifier, path, *, save=True, marker=MARK):
    c = Case(identifier, app, '设计实验非幂等修改', marker, root, path)
    if app == 'word':
        c.add('write', 'writeContent', anchor={'kind': 'documentEnd'}, blocks=[paragraph(marker)])
    elif app == 'excel':
        c.structure('shift', 'insertRows', SHEET, start=1, count=1)
        c.region('write', 'writeRange', SHEET, 'A1', values=[[marker]])
    else:
        sid = c.slide('new', 2)
        c.text('write', sid, marker)
    if save:
        c.finish(save=True)
    return Plan(c.request, path)


def primary_action(app):
    return {'word': 'writeContent', 'excel': 'insertRows', 'ppt': 'addSlide'}[app]
