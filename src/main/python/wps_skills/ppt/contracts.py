"""Deliberate PPT contracts; candidate capabilities are not production discovery."""

import re
import unicodedata
from wps_skills.core.action_session import ActionContract, ApplicationContractSet


def obj(properties, required=None, **extra):
    return dict(type='object', properties=properties,
                required=tuple(properties) if required is None else required,
                additionalProperties=False, **extra)


def text(maximum=32768, **extra):
    return dict(type='string', maxLength=maximum, **extra)


def integer(minimum=0, maximum=2147483647):
    return dict(type='integer', minimum=minimum, maximum=maximum)


def nullable(schema):
    return {'oneOf': ({'type': 'null'}, schema)}


def _pptx_path(value):
    if not value.lower().endswith('.pptx') or any(unicodedata.category(c) in {'Cc', 'Cs'} for c in value):
        return False
    if re.match(r'^[A-Za-z]:[\\/]', value):
        tail = value[3:]
    elif re.match(r'^\\\\[^\\/]+[\\/][^\\/]+[\\/]', value):
        tail = value[2:]
    else:
        return False
    return all(part and not part.endswith((' ', '.')) and not re.search(r'[<>:"|?*]', part)
               and not re.match(r'^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)', part, re.I)
               for part in re.split(r'[\\/]', tail))


ID = integer(1)
TOKEN = text(128, minLength=1)
NUMBER = {'type': 'number'}
BOOL = {'type': 'boolean'}
STATE = obj({'persistenceState': text(enum=('unsaved', 'saved', 'modified')), 'readOnly': BOOL})
ARTIFACT = obj({'path': text(format='absolutePptxPath'), 'format': {'const': 'pptx'}, 'sizeBytes': integer(1, 2**63-1)})
FONT = obj({'latinName': nullable(text(256)), 'eastAsianName': nullable(text(256)), 'size': nullable(NUMBER), 'bold': nullable(BOOL),
            'italic': nullable(BOOL), 'color': nullable(integer())})
SHAPE = obj({'id': ID, 'name': text(), 'type': integer(), 'left': NUMBER, 'top': NUMBER,
             'width': NUMBER, 'height': NUMBER, 'rotation': NUMBER,
             'text': nullable(text(10000)), 'font': nullable(FONT)})
SLIDE = obj({'id': ID, 'index': ID, 'name': text(), 'layout': integer(), 'shapeCount': integer()})
SLIDE_SNAPSHOT = obj({'slide': SLIDE, 'shapes': {'type': 'array', 'maxItems': 100, 'items': SHAPE}, 'token': TOKEN})
SLIDES = obj({'slides': {'type': 'array', 'maxItems': 200, 'items': SLIDE}, 'token': TOKEN})
GEOMETRY = {'left': {'type': 'number', 'minimum': -10000, 'maximum': 10000},
            'top': {'type': 'number', 'minimum': -10000, 'maximum': 10000},
            'width': {'type': 'number', 'minimum': 1, 'maximum': 10000},
            'height': {'type': 'number', 'minimum': 1, 'maximum': 10000}}
STYLE = obj({'latinName': text(128, minLength=1), 'eastAsianName': text(128, minLength=1), 'size': {'type': 'number', 'minimum': 1, 'maximum': 400},
             'bold': BOOL, 'italic': BOOL, 'color': integer(0, 16777215)}, required=(), minProperties=1)
SLIDE_READ = {'slideId': ID}
SLIDE_EDIT = dict(SLIDE_READ, expectedToken=TOKEN)
SHAPE_EDIT = dict(SLIDE_EDIT, shapeId=ID)
ERRORS = ('PERSISTENCE_LOCATOR_REQUIRED', 'OUTPUT_ALREADY_EXISTS', 'OUTPUT_MATCHES_BOUND_DOCUMENT', 'OUTPUT_PARENT_NOT_FOUND', 'OUTPUT_PATH_INVALID', 'OUTPUT_IN_USE', 'OUTPUT_ACCESS_DENIED', 'DOCUMENT_CHANGED_DURING_ACTION', 'OUTPUT_WRITE_FAILED', 'INVALID_PARAMS', 'SESSION_APP_MISMATCH', 'SESSION_DOCUMENT_NOT_BOUND', 'SESSION_DOCUMENT_ALREADY_BOUND',
          'DOCUMENT_NOT_FOUND', 'DOCUMENT_ACCESS_DENIED', 'DOCUMENT_OPEN_FAILED', 'DOCUMENT_BINDING_UNAVAILABLE',
          'DOCUMENT_LEASE_CONFLICT', 'DOCUMENT_QUARANTINED', 'DOCUMENT_CLOSED', 'DOCUMENT_READ_ONLY',
          'PPT_CAPABILITY_UNAVAILABLE', 'SLIDE_NOT_FOUND', 'SHAPE_NOT_FOUND', 'CONTENT_LIMIT_EXCEEDED',
          'TEXT_UNSUPPORTED', 'STALE_CONTENT', 'PPT_READ_FAILED', 'PPT_WRITE_FAILED',
          'PPT_VERIFICATION_FAILED', 'OUTPUT_VERIFICATION_FAILED', 'RESPONSE_LOST', 'IMAGE_NOT_FOUND', 'IMAGE_UNSUPPORTED', 'TABLE_UNSUPPORTED', 'NOTES_UNSUPPORTED', 'SHAPE_UNSUPPORTED')


def _result_error(name, params, result):
    if name == 'createPresentation' and result['documentState']['persistenceState'] != 'unsaved':
        return 'creation must observe an unsaved document'
    if name in {'saveAs', 'exportPdf', 'exportSlideImage'}:
        if result['artifact']['path'] != params['outputPath'] or result['replacedExisting']:
            return 'output must match the absent authorized destination'
        if name == 'saveAs':
            if result['documentState']['persistenceState'] != 'saved':
                return 'saveAs must observe saved state'
        elif result['documentStateBefore'] != result['documentStateAfter']:
            return 'export must preserve document state'
    if name == 'openPresentation' and result['artifact']['path'] != params['path']:
        return 'opened artifact differs from the authorized path'
    if name == 'save' and result['documentState']['persistenceState'] != 'saved':
        return 'save must observe saved state'
    if 'shapes' in result:
        if result['slide']['id'] != params['slideId']:
            return 'result refers to another slide'
        shapes = result['shapes']
        if len(shapes) != result['slide']['shapeCount'] or len({s['id'] for s in shapes}) != len(shapes):
            return 'shape snapshot is incomplete or contains duplicate IDs'
        shape = next((s for s in shapes if s['id'] == params.get('shapeId')), None)
        if name == 'deleteShape':
            if shape is not None:
                return 'deleted shape is still present'
        elif 'shapeId' in params:
            if shape is None:
                return 'modified shape is absent'
            if name == 'setShapeText' and shape['text'] != params['text']:
                return 'text readback differs from requested literal text'
            if name == 'formatText' and (shape['font'] is None or any(shape['font'][k] != v for k, v in params['format'].items())):
                return 'font readback differs from the requested patch'
            if name == 'setShapeGeometry' and any(abs(shape[k] - v) > 0.1 for k, v in params['geometry'].items()):
                return 'geometry readback differs from requested points'
    if 'slides' in result:
        slides = result['slides']
        if len({s['id'] for s in slides}) != len(slides) or [s['index'] for s in slides] != list(range(1, len(slides)+1)):
            return 'slide order is incomplete or has duplicate IDs'
        if name == 'deleteSlide' and any(s['id'] == params['slideId'] for s in slides):
            return 'deleted slide is still present'
        if name == 'moveSlide' and not any(s['id'] == params['slideId'] and s['index'] == params['position'] for s in slides):
            return 'slide did not reach the requested position'
    return _common_result_error(name, params, result)


def contract(name, purpose, category, params, result, example, verification, *, risk='read', role='required', constraint=''):
    return ActionContract(name=name, purpose=purpose, category=category, binding_role=role, risk=risk,
        parameters=params, result=result, stable_errors=ERRORS,
        prerequisites=('Explicit existing-file or new-document intent; UNBOUND Session.' if role == 'establish'
                       else 'One exact live Presentation is bound to this PPT Session.',),
        constraints=(constraint, 'No active-document or selection targeting. No implicit save or automatic replay.'),
        verification=verification, examples=({'params': example},),
        parameter_validator=lambda params: _common_params_error(name, params),
        semantic_validator=lambda params, result: _result_error(name, params, result))


SHAPE['properties']['zOrder'] = ID
SHAPE['properties']['autoShapeType'] = nullable(integer())

_TARGET_CONTRACTS = (
    contract('openPresentation', 'Open or reuse one exact existing .pptx presentation.', 'presentation',
        obj({'path': text(format='absolutePptxPath')}), obj({'artifact': ARTIFACT, 'documentState': STATE}),
        {'path': r'C:\work\slides.pptx'}, 'Verify authorized artifact path and existing dirty state.', role='establish',
        constraint='Only ordinary .pptx files; missing files never become creation requests.'),
    contract('getPresentationInfo', 'Read bound presentation size, saved state and its own window.', 'presentation', obj({}),
        obj({'name': text(), 'documentState': STATE, 'slideCount': integer(), 'width': NUMBER, 'height': NUMBER,
             'window': nullable(obj({'hwnd': ID, 'processId': ID}))}), {},
        'Compare slide dimensions in points and saved state; a null window does not prove desktop visibility.',
        constraint='Window information comes only from the bound presentation.'),
    contract('listSlides', 'Read ordered slide IDs and a structure observation token.', 'slides', obj({}), SLIDES, {},
        'Check ordered IDs and indices; token covers ordered slide summaries.', constraint='At most 200 slides.'),
    contract('getSlideInfo', 'Read top-level shapes, text and aggregate font properties of a slide.', 'slides', obj(SLIDE_READ),
        SLIDE_SNAPSHOT, {'slideId': 256}, 'Check shape IDs, text and geometry. Mixed font properties are null.',
        constraint='At most 100 top-level shapes and 10000 UTF-16 units per shape; token covers returned fields, not notes, animations, groups or individual text runs.'),
    contract('addSlide', 'Insert a blank slide at a 1-based position.', 'slides', obj({'position': integer(1,200), 'expectedToken': TOKEN}),
        SLIDES, {'position': 1, 'expectedToken': 'list-token'}, 'Confirm one new slide at the requested position.', risk='write',
        constraint='Use listSlides token. Position is at most current count + 1; blank layout only.'),
    contract('duplicateSlide', 'Duplicate one slide immediately after itself.', 'slides', obj(SLIDE_EDIT), SLIDES,
        {'slideId':256,'expectedToken':'slide-token'}, 'Confirm one new slide after the source and inspect its contents.', risk='write',
        constraint='Use getSlideInfo token for the source. WPS performs the native duplicate.'),
    contract('moveSlide', 'Move one slide to a 1-based final position.', 'slides', obj(dict(SLIDE_EDIT, position=integer(1,200))), SLIDES,
        {'slideId':256,'position':2,'expectedToken':'list-token'}, 'Confirm source ID appears at the requested final index.', risk='write',
        constraint='Use listSlides token; move retains the stable SlideID.'),
    contract('deleteSlide', 'Delete one explicitly identified slide.', 'slides', obj(SLIDE_EDIT), SLIDES,
        {'slideId':256,'expectedToken':'slide-token'}, 'Confirm the source ID is absent and remaining order is correct.', risk='write',
        constraint='Use getSlideInfo token. Deletion includes notes, animations and all slide content; obtain user intent for deleting the whole slide.'),
    contract('addTextBox', 'Add a horizontal text box with literal text and explicit dimensions.', 'shapes',
        obj(dict(SLIDE_EDIT, text=text(10000, format='pptText'), **GEOMETRY)), SLIDE_SNAPSHOT,
        dict(slideId=256, expectedToken='slide-token', text='Hello', left=40, top=40, width=600, height=100),
        'Confirm one new text box with the requested text and dimensions.', risk='write', constraint='Use getSlideInfo token. Coordinates are points.'),
    contract('addShape', 'Add a rectangle or ellipse at explicit dimensions.', 'shapes',
        obj(dict(SLIDE_EDIT, kind=text(enum=('rectangle','ellipse')), **GEOMETRY)), SLIDE_SNAPSHOT,
        dict(slideId=256,expectedToken='slide-token',kind='rectangle',left=40,top=160,width=200,height=80),
        'Confirm one new shape of the requested kind and dimensions.', risk='write', constraint='Use getSlideInfo token; only two admitted native AutoShape types.'),
    contract('setShapeText', 'Replace all text in one text-capable top-level shape.', 'text',
        obj(dict(SHAPE_EDIT,text=text(10000,format='pptText'))), SLIDE_SNAPSHOT,
        dict(slideId=256,shapeId=2,expectedToken='slide-token',text='Updated'),
        'Compare normalized LF text exactly; inspect formatting after full replacement.', risk='write',
        constraint='Use getSlideInfo token. Full replacement may reset existing rich formatting; no groups, tables, charts or selection.'),
    contract('formatText', 'Apply a uniform font patch to the complete shape text.', 'text',
        obj(dict(SHAPE_EDIT,format=STYLE)), SLIDE_SNAPSHOT,
        dict(slideId=256,shapeId=2,expectedToken='slide-token',format={'size':32,'bold':True}),
        'Compare every requested aggregate font property with native readback.', risk='write',
        constraint='Use getSlideInfo token. Latin and East Asian font names are separate properties. Font color is native OLE RGB integer (red + green*256 + blue*65536); formatting covers all text.'),
    contract('setShapeGeometry', 'Move or resize one top-level shape in points.', 'shapes',
        obj(dict(SHAPE_EDIT,geometry=obj(GEOMETRY,required=(),minProperties=1))), SLIDE_SNAPSHOT,
        dict(slideId=256,shapeId=2,expectedToken='slide-token',geometry={'left':60}),
        'Compare requested coordinates and sizes within 0.1 point.', risk='write',
        constraint='Use getSlideInfo token. Native aspect-ratio behavior is preserved; inspect all dimensions afterwards.'),
    contract('deleteShape', 'Delete one top-level shape by its observed ID.', 'shapes', obj(SHAPE_EDIT), SLIDE_SNAPSHOT,
        dict(slideId=256,shapeId=2,expectedToken='slide-token'), 'Confirm the shape ID is absent.', risk='write',
        constraint='Use getSlideInfo token; deletion removes all content of the specified shape.'),
    contract('save', 'Save the bound existing presentation to its authorized backing file.', 'persistence', obj({}),
        obj({'artifact':ARTIFACT,'documentState':STATE}), {}, 'Verify saved state, nonempty .pptx artifact and continuous exact binding.',
        risk='write',constraint='Use saveAs for first save or a new path. Retain locator and all file-identity fences through native replacement.'),
)
# Common editing contracts use separate observations for style, settings, notes
# and table text, so their tokens cover the properties they actually mutate.
COLOR = integer(0, 16777215)
PARAGRAPH = obj({'alignment': integer(-2,7), 'spaceBefore': NUMBER, 'spaceAfter': NUMBER, 'spaceBeforeInLines': nullable(BOOL), 'spaceAfterInLines': nullable(BOOL),
                 'bulletVisible': nullable(BOOL), 'bulletType': integer(-2,3)})
TEXTBOX = obj({'marginLeft': NUMBER, 'marginRight': NUMBER, 'marginTop': NUMBER, 'marginBottom': NUMBER,
              'verticalAnchor': integer(-2,5), 'wordWrap': nullable(BOOL)})
APPEARANCE = obj({'fillVisible': nullable(BOOL), 'fillType': integer(-2,6), 'fillColor': nullable(COLOR),
                  'fillTransparency': NUMBER, 'lineVisible': nullable(BOOL), 'lineColor': nullable(COLOR), 'lineWidth': NUMBER})
SHAPE_STYLE = obj({'slideId':ID, 'shape':SHAPE, 'appearance':APPEARANCE, 'textBox':nullable(TEXTBOX),
                   'paragraphs':{'type':'array','maxItems':100,'items':PARAGRAPH}, 'token':TOKEN})
APPEARANCE_PATCH = obj({'fillVisible':BOOL,'fillColor':COLOR,
    'fillTransparency':{'type':'number','minimum':0,'maximum':1}, 'lineVisible':BOOL,'lineColor':COLOR,
    'lineWidth':{'type':'number','minimum':0.1,'maximum':20}},required=(),minProperties=1)
PARAGRAPH_PATCH = obj({'alignment':text(enum=('left','center','right','justify')), 'bulletVisible':BOOL,
    'spaceBefore':{'type':'number','minimum':0,'maximum':200}, 'spaceAfter':{'type':'number','minimum':0,'maximum':200}},required=(),minProperties=1)
TEXTBOX_PATCH = obj({**{k:{'type':'number','minimum':0,'maximum':200} for k in ('marginLeft','marginRight','marginTop','marginBottom')},
    'verticalAnchor':text(enum=('top','middle','bottom')), 'wordWrap':BOOL},required=(),minProperties=1)
SETTINGS = obj({'slideId':ID,'name':text(128), 'hidden':BOOL,'followMasterBackground':BOOL,
                'backgroundType':integer(-2,6),'backgroundColor':nullable(COLOR),'token':TOKEN})
SETTINGS_PATCH = obj({'name':text(128,minLength=1,format='pptName'),'hidden':BOOL,
                      'followMasterBackground':BOOL,'backgroundColor':COLOR},required=(),minProperties=1)
NOTES = obj({'slideId':ID,'text':text(10000,format='pptText'),'token':TOKEN})
SHAPE_IDS = {'type':'array','minItems':2,'maxItems':20,'uniqueItems':True,'items':ID}
TABLE_VALUES = {'type':'array','minItems':1,'maxItems':20,
    'items':{'type':'array','minItems':1,'maxItems':10,'items':text(2000,format='pptText')},
    'x-rectangular':True,'x-maxCells':100,'x-maxUtf16Length':20000}
TABLE_SNAPSHOT = obj({'slideId':ID,'shapeId':ID,'rows':integer(1,20),'columns':integer(1,10),'values':TABLE_VALUES,'token':TOKEN})
MATCHES = obj({'slideId':ID, 'matches':{'type':'array','maxItems':100,'items':obj({'shapeId':ID,'start':integer(),'length':ID,'text':text(10000)})},'token':TOKEN})


def _common_params_error(name, params):
    if name == 'formatShape':
        patch=params['format']
        if patch.get('lineVisible') is False and 'lineColor' in patch:
            return 'lineColor enables the border and cannot be combined with lineVisible=false'
        if patch.get('fillVisible') is False and any(k in patch for k in ('fillColor','fillTransparency')):
            return 'fill color/transparency enables the fill and cannot be combined with fillVisible=false'
    if name == 'setSlideSettings' and params['settings'].get('followMasterBackground') and 'backgroundColor' in params['settings']:
        return 'an explicit background color cannot simultaneously inherit the master background'
    if name == 'distributeShapes' and len(params['shapeIds']) < 3:
        return 'distribution requires at least three distinct shapes'
    return None


def _common_result_error(name, params, result):
    if 'slideId' in result and result['slideId'] != params['slideId']:
        return 'result refers to another slide'
    if 'shape' in result and result['shape']['id'] != params['shapeId']:
        return 'style result refers to another shape'
    if name == 'formatShape':
        for k,v in params['format'].items():
            actual=result['appearance'][k]
            if isinstance(v,(int,float)) and not isinstance(v,bool):
                if actual is None or abs(actual-v)>0.01: return 'shape style readback differs from requested patch'
            elif actual != v: return 'shape style readback differs from requested patch'
        if 'lineColor' in params['format'] and result['appearance']['lineVisible'] is not True:
            return 'setting line color must enable the border'
        if any(k in params['format'] for k in ('fillColor','fillTransparency')) and result['appearance']['fillVisible'] is not True:
            return 'setting fill color/transparency must enable the fill'
        if 'fillColor' in params['format'] and result['appearance']['fillType'] != 1:
            return 'explicit fill color requires a solid fill'
    if name == 'formatParagraph':
        if not result['paragraphs']: return 'paragraph formatting requires observed paragraphs'
        for p in result['paragraphs']:
            for k,v in params['format'].items():
                if k=='alignment': v={'left':1,'center':2,'right':3,'justify':4}[v]
                if p[k] != v: return 'paragraph readback differs from requested patch'
                if k in {'spaceBefore','spaceAfter'} and p[k+'InLines'] is not False: return 'paragraph spacing must be in points'
    if name == 'setTextBoxLayout':
        if result['textBox'] is None: return 'shape has no text box layout'
        for k,v in params['layout'].items():
            if k=='verticalAnchor':v={'top':1,'middle':3,'bottom':4}[v]
            actual=result['textBox'][k]
            if isinstance(v,bool):
                if actual != v:return 'text box layout readback differs'
            elif actual is None or abs(actual-v)>0.1:return 'text box layout readback differs'
    if name == 'setSlideSettings':
        for k,v in params['settings'].items():
            if result[k] != v:return 'slide setting readback differs'
        if 'backgroundColor' in params['settings'] and (result['followMasterBackground'] or result['backgroundType']!=1):
            return 'explicit background color requires an independent solid background'
    if name == 'setSlideNotes' and result['text'] != params['text']:
        return 'notes text readback differs'
    if name == 'renameShape':
        shape=next((s for s in result['shapes'] if s['id']==params['shapeId']),None)
        if shape is None or shape['name'] != params['name']:return 'shape name readback differs'
    if name == 'setShapeOrder':
        shape=next((s for s in result['shapes'] if s['id']==params['shapeId']),None)
        if shape is None:return 'ordered shape is absent'
        if params['position']=='front' and shape.get('zOrder')!=len(result['shapes']):return 'shape did not move to front'
        if params['position']=='back' and shape.get('zOrder')!=1:return 'shape did not move to back'
    if name in {'alignShapes','distributeShapes'}:
        shapes=[s for s in result['shapes'] if s['id'] in params['shapeIds']]
        if len(shapes)!=len(params['shapeIds']):return 'arranged shapes are absent'
        if name=='alignShapes':
            key=params['alignment'];axis='left' if key in {'left','center','right'} else 'top'
            size='width' if axis=='left' else 'height';factor=0 if key in {'left','top'} else 0.5 if key in {'center','middle'} else 1
            coords=[s[axis]+factor*s[size] for s in shapes]
            if max(coords)-min(coords)>0.1:return 'shape alignment was not observed'
        else:
            axis,size=('left','width') if params['direction']=='horizontal' else ('top','height')
            shapes.sort(key=lambda s:(s[axis],s['id']))
            gaps=[b[axis]-a[axis]-a[size] for a,b in zip(shapes,shapes[1:])]
            if max(gaps)-min(gaps)>0.2:return 'equal shape gaps were not observed'
    if name=='findText':
        if any(m['text']!=params['text'] or m['length']!=len(params['text'].encode('utf-16-le'))//2 for m in result['matches']):return 'find result does not match the literal query'
        keys=[(m['shapeId'],m['start']) for m in result['matches']]
        if len(set(keys))!=len(keys):return 'duplicate text matches'
    if name in {'addTable','readTable','writeTable'}:
        values=result['values']
        if result['rows']!=len(values) or any(len(row)!=result['columns'] for row in values):return 'table matrix dimensions differ from observed rows/columns'
        if 'shapeId' in params and result['shapeId']!=params['shapeId']:return 'table result refers to another shape'
        if name in {'addTable','writeTable'} and tuple(tuple(row) for row in values)!=tuple(tuple(row) for row in params['values']):return 'table text readback differs from requested matrix'
    return None


_COMMON_TARGET_CONTRACTS = (
    contract('getShapeStyle','Read shape appearance, text box layout and individual paragraph properties.','format',
        obj({'slideId':ID,'shapeId':ID}),SHAPE_STYLE,{'slideId':256,'shapeId':2},
        'Inspect actual fill, border, margins and paragraph values; null means mixed or unavailable.',
        constraint='At most 100 paragraphs; token covers this shape summary, appearance, frame and paragraph fields.'),
    contract('formatShape','Apply a solid fill, transparency and border patch.','format',
        obj(dict(SHAPE_EDIT,format=APPEARANCE_PATCH)),SHAPE_STYLE,dict(slideId=256,shapeId=2,expectedToken='style-token',format={'fillColor':255}),
        'Compare every requested appearance property; explicit fillColor must produce solid fill.',risk='write',constraint='Use getShapeStyle token; colors are OLE RGB integers. Setting fillColor replaces a gradient or picture fill. Fill color/transparency enables fill; lineColor enables the border. Do not combine these with explicit hidden visibility.'),
    contract('formatParagraph','Set uniform paragraph alignment, spacing in points and bullet visibility.','text',
        obj(dict(SHAPE_EDIT,format=PARAGRAPH_PATCH)),SHAPE_STYLE,dict(slideId=256,shapeId=2,expectedToken='style-token',format={'alignment':'center'}),
        'Verify every paragraph matches each requested property.',risk='write',constraint='Use getShapeStyle token; applies to all paragraphs. Only bullet visibility is changed; WPS chooses the native bullet style when enabling bullets.'),
    contract('setTextBoxLayout','Set margins, word wrap and vertical text alignment.','text',
        obj(dict(SHAPE_EDIT,layout=TEXTBOX_PATCH)),SHAPE_STYLE,dict(slideId=256,shapeId=2,expectedToken='style-token',layout={'verticalAnchor':'middle'}),
        'Compare requested margins within 0.1 point and exact wrapping/alignment.',risk='write',constraint='Use getShapeStyle token. Native AutoSize and shape geometry behavior are preserved; inspect shape dimensions afterwards.'),
    contract('renameShape','Rename one top-level shape without changing its ID.','shapes',
        obj(dict(SHAPE_EDIT,name=text(128,minLength=1,format='pptName'))),SLIDE_SNAPSHOT,dict(slideId=256,shapeId=2,expectedToken='slide-token',name='Title'),
        'Verify exact shape name and stable ID.',risk='write',constraint='Use getSlideInfo token; names never become document or shape execution addresses.'),
    contract('setShapeOrder','Bring a shape to front or send it to back.','shapes',
        obj(dict(SHAPE_EDIT,position=text(enum=('front','back')))),SLIDE_SNAPSHOT,dict(slideId=256,shapeId=2,expectedToken='slide-token',position='front'),
        'Verify zOrder equals the shape count for front, or 1 for back.',risk='write',constraint='Use getSlideInfo token; top-level stacking order only.'),
    contract('alignShapes','Align selected shape bounds within their collective bounding box.','shapes',
        obj(dict(SLIDE_EDIT,shapeIds=SHAPE_IDS,alignment=text(enum=('left','center','right','top','middle','bottom')))),SLIDE_SNAPSHOT,
        dict(slideId=256,shapeIds=[2,3],expectedToken='slide-token',alignment='left'),
        'Verify requested edges or centers agree within 0.1 point.',risk='write',constraint='Use getSlideInfo token. 2–20 unrotated top-level shapes; no UI selection, slide-relative alignment or size changes.'),
    contract('distributeShapes','Distribute unrotated shapes with equal horizontal or vertical gaps.','shapes',
        obj(dict(SLIDE_EDIT,shapeIds=SHAPE_IDS,direction=text(enum=('horizontal','vertical')))),SLIDE_SNAPSHOT,
        dict(slideId=256,shapeIds=[2,3,4],expectedToken='slide-token',direction='horizontal'),
        'Verify equal edge-to-edge gaps within 0.2 point and unchanged outer bounds.',risk='write',constraint='Use getSlideInfo token. 3–20 unrotated top-level shapes, with nonnegative available gaps; outer shapes stay fixed.'),
    contract('getSlideSettings','Read slide name, slideshow visibility and background settings.','slides',obj(SLIDE_READ),SETTINGS,{'slideId':256},
        'Inspect hidden state, background inheritance, fill type and color.',constraint='Token covers only returned slide setting fields.'),
    contract('setSlideSettings','Rename a slide, hide it in slideshow or set its background.','slides',
        obj(dict(SLIDE_EDIT,settings=SETTINGS_PATCH)),SETTINGS,dict(slideId=256,expectedToken='settings-token',settings={'hidden':True}),
        'Compare all requested settings; explicit background color disables master inheritance and sets solid fill.',risk='write',constraint='Use getSlideSettings token. Hiding affects slideshow inclusion, not editing visibility.'),
    contract('getSlideNotes','Read the slide notes body placeholder.','notes',obj(SLIDE_READ),NOTES,{'slideId':256},
        'Inspect normalized LF notes text.',constraint='Exactly one body placeholder is required. At most 10000 UTF-16 units; no date/footer/slide-image placeholders.'),
    contract('setSlideNotes','Replace the complete notes body text.','notes',obj(dict(SLIDE_EDIT,text=text(10000,format='pptText'))),NOTES,
        dict(slideId=256,expectedToken='notes-token',text='Speaker notes'),
        'Verify exact normalized notes text.',risk='write',constraint='Use getSlideNotes token. Full replacement can reset notes rich formatting; other notes placeholders are untouched.'),
    contract('findText','Find literal case-sensitive text in a slide’s top-level text shapes.','text',
        obj(dict(SLIDE_READ,text=text(10000,minLength=1,format='pptText'))),MATCHES,dict(slideId=256,text='Hello'),
        'Inspect matched shape IDs and zero-based UTF-16 offsets; results never address a document.',constraint='At most 100 nonoverlapping matches. No regex, groups, tables, notes or cross-slide search; token is the getSlideInfo observation.'),
    contract('replaceText','Replace every literal case-sensitive occurrence in one shape.','text',
        obj(dict(SHAPE_EDIT,find=text(10000,minLength=1,format='pptText'),replacement=text(10000,format='pptText'))),
        obj(dict(SLIDE_SNAPSHOT['properties'],replacements=integer(0,100))),
        dict(slideId=256,shapeId=2,expectedToken='slide-token',find='Hello',replacement='Hi'),
        'Compare complete expected text and replacement count after native text-range edits.',risk='write',constraint='Use getSlideInfo token. At most 100 nonoverlapping matches; edits run from end to start, preserving text outside replaced spans. Result text must fit 10000 UTF-16 units.'),
    contract('addImage','Embed one existing PNG or JPEG at explicit slide coordinates.','images',
        obj(dict(SLIDE_EDIT,path=text(format='absolutePptImagePath'),**GEOMETRY)),SLIDE_SNAPSHOT,
        dict(slideId=256,expectedToken='slide-token',path=r'C:\work\picture.png',left=40,top=40,width=400,height=300),
        'Verify exactly one new picture shape and requested dimensions; image bytes must be embedded on save.',risk='write',constraint='Use getSlideInfo token. Absolute existing host path, at most 20 MiB and 40 million pixels; embedded, never linked. Explicit dimensions may change aspect ratio.'),
    contract('addTable','Insert a native table containing a rectangular literal-text matrix.','tables',
        obj(dict(SLIDE_EDIT,values=TABLE_VALUES,**GEOMETRY)),TABLE_SNAPSHOT,
        dict(slideId=256,expectedToken='slide-token',values=[['Item','Value'],['A','42']],left=40,top=160,width=600,height=200),
        'Read each native cell and compare the complete matrix and dimensions.',risk='write',constraint='Use getSlideInfo token. At most 20 rows, 10 columns, 100 cells and 20000 total UTF-16 units. Literal text only.'),
    contract('readTable','Read all literal cell text from one top-level native table.','tables',
        obj({'slideId':ID,'shapeId':ID}),TABLE_SNAPSHOT,dict(slideId=256,shapeId=3),
        'Check row/column counts and all cell text.',constraint='Same table bounds as addTable; token covers table ID, dimensions and cell text, not styling or merged-cell geometry.'),
    contract('writeTable','Replace all cell text of one existing table without resizing.','tables',
        obj(dict(SHAPE_EDIT,values=TABLE_VALUES)),TABLE_SNAPSHOT,
        dict(slideId=256,shapeId=3,expectedToken='table-token',values=[['A','B']]),
        'Compare every cell with the complete requested matrix.',risk='write',constraint='Use readTable token. Matrix must exactly match existing rows/columns. Cell rich formatting may reset; merged cells may reject or yield an uncertain partial result, never auto-retry.'),
)
_TARGET_CONTRACTS += _COMMON_TARGET_CONTRACTS


# Persistence readback is bounded to 200 slides, each at most 100 top-level shapes.
_TARGET_CONTRACTS += (
    contract('createPresentation', 'Create one blank unsaved document; no file path.', 'persistence', obj({}), obj({'documentState': STATE}), {}, 'Verify exact binding, artifact format and unchanged observed content/state.', constraint='', role='establish', risk='write'),
    contract('saveAs', 'First-save or save the same live document to an absent destination.', 'persistence', obj({'outputPath': text(format='absolutePptxPath'), 'overwritePolicy': {'const': 'failIfExists'}}), obj({'artifact': ARTIFACT, 'documentState': STATE, 'replacedExisting': {'const': False}}), {'outputPath': 'C:/work/new.pptx', 'overwritePolicy': 'failIfExists'}, 'Verify exact binding, artifact format and unchanged observed content/state.', constraint='At most 200 slides and 100 top-level shapes per slide. Never replaces a preexisting destination. Retain old and new locator and identity fences until Session cleanup.', role='required', risk='write'),
    contract('exportPdf', 'Export to an absent destination without saving the document.', 'persistence', obj({'outputPath': text(format='absolutePdfPath'), 'overwritePolicy': {'const': 'failIfExists'}}), obj({'artifact': obj({'path': text(format='absolutePdfPath'), 'format': {'const': 'pdf'}, 'sizeBytes': {'type': 'integer', 'minimum': 1}}), 'documentStateBefore': STATE, 'documentStateAfter': STATE, 'replacedExisting': {'const': False}}), {'outputPath': 'C:/work/output.pdf', 'overwritePolicy': 'failIfExists'}, 'Verify exact binding, artifact format and unchanged observed content/state.', constraint='Native PDF export defaults; PNG uses explicit pixels. At most 200 slides and 100 top-level shapes per slide. Verify exported format and unchanged bounded observations; no save or retargeting.', role='required', risk='write'),
    contract('exportSlideImage', 'Export to an absent destination without saving the document.', 'persistence', obj({'outputPath': text(format='absolutePngPath'), 'overwritePolicy': {'const': 'failIfExists'}, 'slideId': ID, 'width': integer(1,4096), 'height': integer(1,4096)}), obj({'artifact': obj({'path': text(format='absolutePngPath'), 'format': {'const': 'png'}, 'sizeBytes': {'type': 'integer', 'minimum': 1}}), 'documentStateBefore': STATE, 'documentStateAfter': STATE, 'replacedExisting': {'const': False}}), {'outputPath': 'C:/work/output.png', 'overwritePolicy': 'failIfExists', 'slideId': 256, 'width': 1280, 'height': 720}, 'Verify exact binding, artifact format and unchanged observed content/state.', constraint='Native PDF export defaults; PNG uses explicit pixels. At most 200 slides and 100 top-level shapes per slide. Verify exported format and unchanged bounded observations; no save or retargeting.', role='required', risk='write'),
)

PPT_FORMAT_VALIDATORS = {'absolutePptxPath': _pptx_path,
    'absolutePptImagePath': lambda value: any(value.lower().endswith(ext) and _pptx_path(value[:-len(ext)]+'.pptx') for ext in ('.png','.jpg','.jpeg')),
    'pptName': lambda value: all(unicodedata.category(c) not in {'Cc','Cs'} for c in value),
    'pptText': lambda value: len(value.encode('utf-16-le', errors='surrogatepass')) <= 20000 and '\r' not in value and all(unicodedata.category(c) not in {'Cc','Cs'} or c in '\n\t' for c in value)}
PPT_FORMAT_VALIDATORS.update({
    'absolutePdfPath': lambda value: value.lower().endswith('.pdf') and _pptx_path(value[:-4]+'.pptx'),
    'absolutePngPath': lambda value: value.lower().endswith('.png') and _pptx_path(value[:-4]+'.pptx'),
})
PPT_TARGET_CONTRACT_SET = ApplicationContractSet(application='ppt', contracts=_TARGET_CONTRACTS, format_validators=PPT_FORMAT_VALIDATORS)
# Base slice and common Actions admitted after native acceptance on 2026-09-05.
# See src/test/resources/wps_skills/ppt/type_library/EVIDENCE.md.
_PPT_PRODUCTION_ACTIONS = frozenset({
    'saveAs', 'createPresentation', 'exportPdf', 'exportSlideImage',
    'openPresentation', 'getPresentationInfo', 'listSlides', 'getSlideInfo',
    'addSlide', 'duplicateSlide', 'moveSlide', 'deleteSlide', 'addTextBox',
    'addShape', 'setShapeText', 'formatText', 'setShapeGeometry', 'deleteShape', 'save',
    'getShapeStyle', 'formatShape', 'formatParagraph', 'setTextBoxLayout',
    'renameShape', 'setShapeOrder', 'alignShapes', 'distributeShapes',
    'getSlideSettings', 'setSlideSettings', 'getSlideNotes', 'setSlideNotes',
    'findText', 'replaceText', 'addImage', 'addTable', 'readTable', 'writeTable',
})
PPT_PRODUCTION_CONTRACT_SET = ApplicationContractSet(application='ppt',
    contracts=tuple(c for c in _TARGET_CONTRACTS if c.name in _PPT_PRODUCTION_ACTIONS),
    format_validators=PPT_FORMAT_VALIDATORS)
PPT_PRODUCTION_ACTION_INDEX = PPT_PRODUCTION_CONTRACT_SET.action_index()
