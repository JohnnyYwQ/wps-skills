"""Bounded Excel contracts; live-WPS admission is separate from implementation."""

import re
import unicodedata

from wps_skills.core.action_session import ActionContract, ApplicationContractSet

MAX_CELLS = 1000


def range_bounds(address):
    """Canonical, sheet-local A1 rectangles only; no names or COM expressions."""
    match = re.fullmatch(r"([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?", address)
    if match is None:
        raise ValueError("use an uppercase A1 cell or rectangle without sheet qualifiers")
    def column(letters):
        value = 0
        for letter in letters:
            value = value * 26 + ord(letter) - ord('A') + 1
        return value
    c1, r1 = column(match[1]), int(match[2])
    c2, r2 = column(match[3] or match[1]), int(match[4] or match[2])
    if not (1 <= c1 <= c2 <= 16384 and 1 <= r1 <= r2 <= 1048576):
        raise ValueError("range is reversed or outside worksheet bounds")
    if (r2 - r1 + 1) * (c2 - c1 + 1) > MAX_CELLS:
        raise ValueError("range exceeds the 1000-cell Action limit")
    return r1, c1, r2, c2


def _xlsx_path(value):
    if not value.lower().endswith('.xlsx') or any(unicodedata.category(c) in {'Cc', 'Cs'} for c in value):
        return False
    if re.match(r'^[A-Za-z]:[\\/]', value):
        tail = value[3:]
    elif re.match(r'^\\\\[^\\/]+[\\/][^\\/]+[\\/]', value):
        tail = value[2:]
    else:
        return False
    parts = re.split(r'[\\/]', tail)
    return all(part and not part.endswith((' ', '.')) and not re.search(r'[<>:"|?*]', part)
               and not re.match(r'^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)', part, re.I)
               for part in parts)



_FORMULA_TOKEN = re.compile(r'\s+|"(?:[^"\r\n]|"")*"|(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?|\$?[A-Za-z_][A-Za-z0-9_.$]*|[=+*/^%&<>()\-,:]')
_FORMULA_FUNCTIONS = frozenset({'SUM', 'AVERAGE', 'COUNT', 'MIN', 'MAX', 'IF', 'IFERROR', 'ROUND', 'ABS', 'COUNTIF', 'COUNTIFS', 'SUMIF', 'SUMIFS', 'COUNTA', 'COUNTBLANK',
    'AND', 'OR', 'NOT', 'LEFT', 'RIGHT', 'MID', 'LEN', 'TRIM', 'UPPER', 'LOWER', 'CONCATENATE',
    'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'DATE', 'YEAR', 'MONTH', 'DAY',
    'ROUNDUP', 'ROUNDDOWN', 'TEXT', 'VALUE', 'SUBSTITUTE', 'SEARCH', 'FIND'})


def _formula(value):
    if not value.startswith('='):
        return False
    position = 0
    tokens = []
    for match in _FORMULA_TOKEN.finditer(value):
        if match.start() != position:
            return False
        position = match.end()
        if not match[0].isspace():
            tokens.append(match[0])
    if position != len(value):
        return False
    for index, token in enumerate(tokens):
        if re.match(r'^\$?[A-Za-z_]', token):
            name = token.upper()
            if index + 1 < len(tokens) and tokens[index + 1] == '(':
                if name not in _FORMULA_FUNCTIONS:
                    return False
            elif name not in {'TRUE', 'FALSE'}:
                try:
                    range_bounds(name.replace('$', ''))
                except ValueError:
                    return False
    return True


def obj(properties, required=None, **extra):
    return dict(type='object', properties=properties, required=tuple(properties) if required is None else required,
                additionalProperties=False, **extra)


def string(maximum=4096, **extra):
    return dict(type='string', maxLength=maximum, **extra)


def matrix(items):
    return {'type': 'array', 'minItems': 1, 'maxItems': MAX_CELLS,
            'items': {'type': 'array', 'minItems': 1, 'maxItems': MAX_CELLS, 'items': items},
            'x-rectangular': True, 'x-maxCells': MAX_CELLS, 'x-maxUtf16Length': 262144}


SCALAR = {'oneOf': ({'type': 'null'}, {'type': 'boolean'}, {'type': 'number'}, string(format='excelLiteral'))}
READ_SCALAR = {'oneOf': ({'type': 'null'}, {'type': 'boolean'}, {'type': 'number'}, string(32767))}
TOKEN = string(128, minLength=1)
SHEET = string(31, minLength=1)
ADDRESS = string(32, minLength=1)
STATE = obj({'persistenceState': string(enum=('unsaved', 'saved', 'modified')), 'readOnly': {'type': 'boolean'}})
ARTIFACT = obj({'path': string(32768, format='absoluteXlsxPath'), 'format': {'const': 'xlsx'},
                'sizeBytes': {'type': 'integer', 'minimum': 1}})
CELL = obj({'value': READ_SCALAR, 'formula': {'oneOf': ({'type': 'null'}, string(8192, minLength=1))},
            'text': string(32768), 'errorCode': {'oneOf': ({'type': 'null'}, {'type': 'integer'})},
            'numberFormat': string(1024), 'bold': {'type': 'boolean'}})
SNAPSHOT = obj({'sheet': SHEET, 'address': ADDRESS, 'token': TOKEN, 'cells': matrix(CELL)})
REGION = {'sheet': SHEET, 'address': ADDRESS}
STYLE = {
    'italic': {'type': 'boolean'}, 'fontSize': {'type': 'number', 'minimum': 1, 'maximum': 409},
    'fontColor': {'type': 'integer', 'minimum': 0, 'maximum': 16777215},
    'fillColor': {'type': 'integer', 'minimum': 0, 'maximum': 16777215},
    'wrapText': {'type': 'boolean'},
    'horizontalAlignment': string(enum=('general', 'left', 'center', 'right')),
    'verticalAlignment': string(enum=('top', 'center', 'bottom')),
}
CELL['properties'].update(STYLE)
CELL['properties']['horizontalAlignment'] = string(enum=('general', 'left', 'center', 'right', 'other'))
CELL['properties']['verticalAlignment'] = string(enum=('top', 'center', 'bottom', 'other'))
CELL['properties'].update({'merged': {'type': 'boolean'}, 'rowHidden': {'type': 'boolean'},
    'columnHidden': {'type': 'boolean'}, 'rowHeight': {'type': 'number', 'minimum': 0},
    'columnWidth': {'type': 'number', 'minimum': 0}})
PATCH = obj(dict(STYLE, numberFormat=string(128, minLength=1), bold={'type': 'boolean'}), required=(), minProperties=1)
ERRORS = ('PERSISTENCE_LOCATOR_REQUIRED', 'OUTPUT_ALREADY_EXISTS', 'OUTPUT_MATCHES_BOUND_DOCUMENT', 'OUTPUT_PARENT_NOT_FOUND', 'OUTPUT_PATH_INVALID', 'OUTPUT_IN_USE', 'OUTPUT_ACCESS_DENIED', 'DOCUMENT_CHANGED_DURING_ACTION', 'INVALID_PARAMS', 'UNKNOWN_ACTION', 'SESSION_APP_MISMATCH', 'SESSION_DOCUMENT_NOT_BOUND',
          'SESSION_DOCUMENT_ALREADY_BOUND', 'DOCUMENT_NOT_FOUND', 'DOCUMENT_ACCESS_DENIED',
          'DOCUMENT_OPEN_FAILED', 'DOCUMENT_BINDING_UNAVAILABLE', 'DOCUMENT_LEASE_CONFLICT',
          'DOCUMENT_QUARANTINED', 'DOCUMENT_CLOSED', 'DOCUMENT_READ_ONLY', 'WORKSHEET_NOT_FOUND',
          'RANGE_UNSUPPORTED', 'STALE_RANGE', 'RANGE_READ_FAILED', 'RANGE_WRITE_FAILED',
          'RANGE_VERIFICATION_FAILED', 'OUTPUT_WRITE_FAILED', 'OUTPUT_VERIFICATION_FAILED',
          'EXCEL_CAPABILITY_UNAVAILABLE', 'RESPONSE_LOST')


def _params_error(action, params):
    for key in ('sheet', 'name', 'targetSheet'):
        if key in params and (any(c in params[key] for c in '[]:*?/\\') or params[key].startswith("'") or params[key].endswith("'") or any(unicodedata.category(c) in {'Cc', 'Cs'} for c in params[key])):
            return 'invalid worksheet name'
    if action in {'insertRows', 'deleteRows', 'insertColumns', 'deleteColumns'}:
        maximum = 1048576 if action.endswith('Rows') else 16384
        if params['start'] + params['count'] - 1 > maximum:
            return 'structural edit exceeds worksheet bounds'
    if 'targetAddress' in params:
        try:
            a, b, c, d = range_bounds(params['address'])
            e, f, g, h = range_bounds(params['targetAddress'])
        except ValueError as exc:
            return str(exc)
        if (c-a, d-b) != (g-e, h-f):
            return 'copy source and destination must have the same shape'
        if params['sheet'] == params['targetSheet'] and not (g < a or e > c or h < b or f > d):
            return 'copy source and destination must not overlap'
    if 'address' in params:
        try:
            r1, c1, r2, c2 = range_bounds(params['address'])
        except ValueError as exc:
            return str(exc)
        if action in {'sortRange', 'filterRange'} and params['column'] > c2 - c1 + 1:
            return 'column is relative to the rectangle and must be inside it'
        if action in {'sortRange', 'filterRange'} and r2 == r1:
            return 'sort and filter require at least two rows'
        field = 'values' if action == 'writeRange' else 'formulas' if action == 'setFormulas' else None
        if field:
            rows = params[field]
            if len(rows) != r2 - r1 + 1 or any(len(row) != c2 - c1 + 1 for row in rows):
                return 'matrix shape must exactly match the requested rectangle'
        if any(c in params['sheet'] for c in '[]:*?/\\'):
            return 'invalid worksheet name'
    return None


def _result_error(action, params, result):
    if action == 'createWorkbook' and result['documentState']['persistenceState'] != 'unsaved':
        return 'creation must observe an unsaved document'
    if action in {'saveAs', 'exportPdf', 'exportSlideImage'}:
        if result['artifact']['path'] != params['outputPath'] or result['replacedExisting']:
            return 'output must match the absent authorized destination'
        if action == 'saveAs':
            if result['documentState']['persistenceState'] != 'saved':
                return 'saveAs must observe saved state'
        elif result['documentStateBefore'] != result['documentStateAfter']:
            return 'export must preserve document state'
    if action == 'openWorkbook' and result['artifact']['path'] != params['path']:
        return 'opened artifact must match the authorized path'
    if action == 'listWorksheets':
        items = result['worksheets']
        if len(items) > params['limit']:
            return 'worksheet page exceeds the requested limit'
        if [s['index'] for s in items] != list(range(params['offset'] + 1, params['offset'] + len(items) + 1)):
            return 'worksheet indices must match the requested page'
        end = params['offset'] + len(items)
        expected = end if end < result['total'] else None
        if result['nextOffset'] != expected or end > result['total'] and items:
            return 'worksheet pagination facts disagree'
    if action in {'getWorksheetInfo', 'addWorksheet', 'renameWorksheet', 'copyWorksheet', 'moveWorksheet', 'insertRows', 'deleteRows', 'insertColumns', 'deleteColumns'}:
        if result['name'] != params.get('name', params['sheet']):
            return 'worksheet result must name the requested worksheet'
        if action == 'moveWorksheet' and result['index'] != params['index']:
            return 'worksheet move position differs from requested index'
    if action == 'deleteWorksheet' and result['deleted'] != params['sheet']:
        return 'deletion result must name the requested worksheet'
    if action == 'findInRange':
        if result['sheet'] != params['sheet'] or result['address'] != params['address']:
            return 'search must describe the requested region'
        top, left, bottom, right = range_bounds(params['address'])
        seen = set()
        for match in result['matches']:
            try:
                row, column, end_row, end_column = range_bounds(match['address'])
            except ValueError as exc:
                return str(exc)
            if row != end_row or column != end_column or not (top <= row <= bottom and left <= column <= right) or match['address'] in seen:
                return 'search matches must be unique cells within the requested region'
            seen.add(match['address'])
            value, wanted = match['text'], params['text']
            if not params['matchCase']:
                value, wanted = value.lower(), wanted.lower()
            if (value != wanted if params['wholeCell'] else wanted not in value):
                return 'search result does not match the literal query'
        return None
    if action == 'copyRange':
        params = dict(params, sheet=params['targetSheet'], address=params['targetAddress'])
    if 'address' not in params:
        return None
    if result['sheet'] != params['sheet'] or result['address'] != params['address']:
        return 'snapshot must describe the requested sheet and rectangle'
    r1, c1, r2, c2 = range_bounds(params['address'])
    cells = result['cells']
    if len(cells) != r2 - r1 + 1 or any(len(row) != c2 - c1 + 1 for row in cells):
        return 'snapshot shape does not match the requested rectangle'
    for row_index, row in enumerate(cells):
        for col_index, cell in enumerate(row):
            if cell['errorCode'] is not None and cell['value'] is not None:
                return 'error cells must carry a null value and a separate error code'
            if action == 'writeRange':
                wanted = params['values'][row_index][col_index]
                if cell['formula'] is not None or cell['errorCode'] is not None or cell['value'] != wanted:
                    return 'observed value differs from literal write'
                if isinstance(wanted, bool) != isinstance(cell['value'], bool):
                    return 'Boolean and numeric values must retain their types'
            if action == 'setFormulas' and cell['formula'] != params['formulas'][row_index][col_index]:
                return 'observed formula differs from requested formula'
            if action == 'clearRange' and params['mode'] in {'contents', 'all'} and (cell['value'] is not None or cell['formula'] is not None):
                return 'clear did not remove cell content'
            if action == 'mergeRange' and not cell.get('merged'):
                return 'merge was not observed'
            if action == 'unmergeRange' and cell.get('merged'):
                return 'unmerge was not observed'
            if action == 'setRowHeight' and abs(cell.get('rowHeight', -1000) - params['height']) > 0.8:
                return 'observed row height differs from requested height'
            if action == 'setColumnWidth' and abs(cell.get('columnWidth', -1000) - params['width']) > 0.2:
                return 'observed column width differs from requested width'
            if action == 'autoFitColumns' and cell.get('columnWidth', 0) <= 0:
                return 'autofit did not produce a visible column'
            if action == 'formatRange' and any(cell[key] != value for key, value in params['format'].items()):
                return 'observed formatting differs from requested patch'
    return None


def contract(name, purpose, category, params, result, example, *, risk='read', role='required'):
    return ActionContract(
        name=name, purpose=purpose, category=category, binding_role=role, risk=risk,
        parameters=params, result=result, stable_errors=ERRORS,
        prerequisites=('Explicit existing-file or new-document intent; UNBOUND Session.' if role == 'establish'
                       else 'One exact live Workbook is bound to this Excel Session.',),
        constraints=('One worksheet-local contiguous A1 rectangle, at most 1000 cells per Action.',
                     'Region writes use readRange tokens; worksheet and structure writes use getWorksheetInfo tokens. No automatic replay.',
                     'No active-window targeting, external workbook references, or implicit save.',
                     'Values are JSON scalars; dates are serial numbers interpreted with date1904 and numberFormat.',
                     'Formula dialect and calculation compatibility require validation on the installed WPS version.',
                     *(('Save As and PDF export verify at most 20 worksheets, each UsedRange at most 1000 cells; only absent destinations with failIfExists. PDF follows current native print settings.',) if name in {'saveAs','exportPdf'} else ())),
        verification='Read back the exact region and compare values, formulas and requested format; inspect errors. Save must prove continuous binding and the saved artifact.',
        examples=({'params': example},),
        parameter_validator=lambda params: _params_error(name, params),
        semantic_validator=lambda params, result: _result_error(name, params, result),
    )


_EXAMPLE_REGION = {'sheet': 'Sheet1', 'address': 'A1:B1'}
_MUTATION_REGION = dict(_EXAMPLE_REGION, expectedToken='read-token')
_TARGET_CONTRACTS = (
    contract('openWorkbook', 'Open or reuse one exact existing .xlsx workbook.', 'workbook',
             obj({'path': string(32768, format='absoluteXlsxPath')}),
             obj({'artifact': ARTIFACT, 'documentState': STATE}), {'path': r'C:\work\book.xlsx'}, role='establish'),
    contract('getWorkbookInfo', 'Inspect bound workbook identity, saved state and date system.', 'workbook',
             obj({}), obj({'name': string(32768, minLength=1), 'documentState': STATE,
                           'date1904': {'type': 'boolean'}, 'worksheetCount': {'type': 'integer', 'minimum': 0},
                           'window': {'oneOf': ({'type': 'null'}, obj({'hwnd': {'type': 'integer', 'minimum': 1}, 'processId': {'type': 'integer', 'minimum': 1}}))}}), {}),
    contract('listWorksheets', 'Read a bounded page of worksheet names and indices.', 'worksheets',
             obj({'offset': {'type': 'integer', 'minimum': 0}, 'limit': {'type': 'integer', 'minimum': 1, 'maximum': 100}}),
             obj({'worksheets': {'type': 'array', 'maxItems': 100, 'items': obj({'name': SHEET, 'index': {'type': 'integer', 'minimum': 1}})},
                  'total': {'type': 'integer', 'minimum': 0}, 'nextOffset': {'oneOf': ({'type': 'null'}, {'type': 'integer', 'minimum': 1})}}),
             {'offset': 0, 'limit': 100}),
    contract('readRange', 'Read values, formulas, displayed text and format with a region token.', 'ranges',
             obj(REGION), SNAPSHOT, _EXAMPLE_REGION),
    contract('writeRange', 'Write an exact matrix of literal values and verify the result.', 'ranges',
             obj(dict(REGION, expectedToken=TOKEN, values=matrix(SCALAR))), SNAPSHOT,
             dict(_MUTATION_REGION, values=[['Item', 42]]), risk='write'),
    contract('setFormulas', 'Write an exact matrix of invariant A1 formulas and read them back.', 'formulas',
             obj(dict(REGION, expectedToken=TOKEN, formulas=matrix(string(8192, minLength=2, format='excelFormula')))), SNAPSHOT,
             dict(_MUTATION_REGION, formulas=[['=1+1', '=A1*2']]), risk='write'),
    contract('calculateRange', 'Recalculate the specified bound-workbook region and inspect results.', 'formulas',
             obj(dict(REGION, expectedToken=TOKEN)), SNAPSHOT, _MUTATION_REGION, risk='write'),
    contract('formatRange', 'Set number, font, color, wrapping and alignment formats, verifying each cell.', 'formatting',
             obj(dict(REGION, expectedToken=TOKEN, format=PATCH)), SNAPSHOT,
             dict(_MUTATION_REGION, format={'bold': True}), risk='write'),
    contract('save', 'Save the existing .xlsx locator and verify persistence without retargeting.', 'persistence',
             obj({}), obj({'artifact': ARTIFACT, 'documentState': obj({'persistenceState': {'const': 'saved'}, 'readOnly': {'const': False}})}),
             {}, risk='write'),
)

SHEET_INFO = obj({'name': SHEET, 'index': {'type': 'integer', 'minimum': 1},
    'visible': {'type': 'boolean'}, 'usedAddress': ADDRESS, 'token': TOKEN})
_SHEET_PARAMS = {'sheet': SHEET, 'expectedToken': TOKEN}
_SHEET_EXAMPLE = {'sheet': 'Sheet1', 'expectedToken': 'worksheet-token'}
_EXTRA_CONTRACTS = [
    contract('getWorksheetInfo', 'Inspect worksheet structure and bounded used content with an edit token.', 'worksheets',
             obj({'sheet': SHEET}), SHEET_INFO, {'sheet': 'Sheet1'}),
]
for name, purpose, fields, example in (
    ('addWorksheet', 'Add a named worksheet after the specified worksheet.', {'name': SHEET}, {'name': 'New'}),
    ('renameWorksheet', 'Rename the exact worksheet and verify its new name.', {'name': SHEET}, {'name': 'Renamed'}),
    ('copyWorksheet', 'Copy the exact worksheet within this workbook under a new name.', {'name': SHEET}, {'name': 'Copy'}),
    ('moveWorksheet', 'Move the worksheet to a one-based position in this workbook.', {'index': {'type': 'integer', 'minimum': 1, 'maximum': 1000}}, {'index': 1}),
    ('deleteWorksheet', 'Delete the observed worksheet; refuse the last visible worksheet.', {}, {}),
    ('insertRows', 'Insert whole rows in a bounded worksheet and verify the shifted cells.', {'start': {'type': 'integer', 'minimum': 1, 'maximum': 1048576}, 'count': {'type': 'integer', 'minimum': 1, 'maximum': 100}}, {'start': 2, 'count': 1}),
    ('deleteRows', 'Delete whole rows in a bounded worksheet and verify remaining cells.', {'start': {'type': 'integer', 'minimum': 1, 'maximum': 1048576}, 'count': {'type': 'integer', 'minimum': 1, 'maximum': 100}}, {'start': 2, 'count': 1}),
    ('insertColumns', 'Insert whole columns in a bounded worksheet.', {'start': {'type': 'integer', 'minimum': 1, 'maximum': 16384}, 'count': {'type': 'integer', 'minimum': 1, 'maximum': 100}}, {'start': 2, 'count': 1}),
    ('deleteColumns', 'Delete whole columns in a bounded worksheet.', {'start': {'type': 'integer', 'minimum': 1, 'maximum': 16384}, 'count': {'type': 'integer', 'minimum': 1, 'maximum': 100}}, {'start': 2, 'count': 1}),
):
    result = obj({'deleted': SHEET}) if name == 'deleteWorksheet' else SHEET_INFO
    _EXTRA_CONTRACTS.append(contract(name, purpose, 'worksheets' if 'Worksheet' in name else 'structure',
        obj(dict(_SHEET_PARAMS, **fields)), result, dict(_SHEET_EXAMPLE, **example), risk='write'))
for name, purpose, fields, example in (
    ('clearRange', 'Clear contents, formatting, or both in the observed region.', {'mode': string(enum=('contents', 'formats', 'all'))}, {'mode': 'contents'}),
    ('mergeRange', 'Merge a rectangle only when all cells except its top-left are empty.', {}, {}),
    ('unmergeRange', 'Unmerge only merged areas wholly contained in the rectangle.', {}, {}),
    ('sortRange', 'Sort constant-value rows by one relative column, with explicit header handling.', {'column': {'type': 'integer', 'minimum': 1, 'maximum': 1000}, 'order': string(enum=('ascending', 'descending')), 'header': {'type': 'boolean'}}, {'column': 1, 'order': 'ascending', 'header': True}),
    ('filterRange', 'Apply one literal equals filter, preserving the header row.', {'column': {'type': 'integer', 'minimum': 1, 'maximum': 1000}, 'value': string(255, format='excelLiteral')}, {'column': 1, 'value': 'Item'}),
    ('clearFilter', 'Remove the worksheet filter only if its range matches the observed rectangle.', {}, {}),
    ('setRowHeight', 'Set row height in points for rows intersecting this region.', {'height': {'type': 'number', 'minimum': 1, 'maximum': 409}}, {'height': 24}),
    ('setColumnWidth', 'Set column width in character units for columns intersecting this region.', {'width': {'type': 'number', 'minimum': 1, 'maximum': 255}}, {'width': 18}),
    ('autoFitColumns', 'Auto-fit columns to the cells in the specified region.', {}, {}),
    ('replaceInRange', 'Replace literal text in string constants; never edit formulas.', {'text': string(4096, minLength=1, format='excelLiteral'), 'replacement': string(4096, format='excelLiteral'), 'matchCase': {'type': 'boolean'}, 'wholeCell': {'type': 'boolean'}}, {'text': 'old', 'replacement': 'new', 'matchCase': True, 'wholeCell': False}),
):
    _EXTRA_CONTRACTS.append(contract(name, purpose, 'ranges', obj(dict(REGION, expectedToken=TOKEN, **fields)),
        SNAPSHOT, dict(_MUTATION_REGION, **dict(example, address='A1:B3')), risk='write'))
_EXTRA_CONTRACTS.extend((
    contract('copyRange', 'Copy literal values to an equally sized nonoverlapping region, validating both tokens.', 'ranges',
        obj(dict(REGION, expectedToken=TOKEN, targetSheet=SHEET, targetAddress=ADDRESS, targetToken=TOKEN)), SNAPSHOT,
        dict(_MUTATION_REGION, targetSheet='Sheet1', targetAddress='A3:B3', targetToken='target-token'), risk='write'),
    contract('findInRange', 'Find literal text in values or formulas within a bounded region.', 'ranges',
        obj(dict(REGION, text=string(4096, minLength=1), matchCase={'type': 'boolean'}, wholeCell={'type': 'boolean'}, lookIn=string(enum=('values', 'formulas')))),
        obj({'sheet': SHEET, 'address': ADDRESS, 'token': TOKEN, 'matches': {'type': 'array', 'maxItems': 1000, 'items': obj({'address': ADDRESS, 'text': string(32768)})}}),
        dict(_EXAMPLE_REGION, text='Item', matchCase=False, wholeCell=False, lookIn='values')),
))
_TARGET_CONTRACTS += tuple(_EXTRA_CONTRACTS)


# Persistence readback is bounded to 20 worksheets, each UsedRange at most 1000 cells.
_TARGET_CONTRACTS += (
    contract('createWorkbook', 'Create one blank unsaved document; no file path.', 'persistence', obj({}), obj({'documentState': STATE}), {}, role='establish', risk='write'),
    contract('saveAs', 'First-save or save the same live document to an absent destination.', 'persistence', obj({'outputPath': string(format='absoluteXlsxPath'), 'overwritePolicy': {'const': 'failIfExists'}}), obj({'artifact': ARTIFACT, 'documentState': STATE, 'replacedExisting': {'const': False}}), {'outputPath': 'C:/work/new.xlsx', 'overwritePolicy': 'failIfExists'}, role='required', risk='write'),
    contract('exportPdf', 'Export to an absent destination without saving the document.', 'persistence', obj({'outputPath': string(format='absolutePdfPath'), 'overwritePolicy': {'const': 'failIfExists'}}), obj({'artifact': obj({'path': string(format='absolutePdfPath'), 'format': {'const': 'pdf'}, 'sizeBytes': {'type': 'integer', 'minimum': 1}}), 'documentStateBefore': STATE, 'documentStateAfter': STATE, 'replacedExisting': {'const': False}}), {'outputPath': 'C:/work/output.pdf', 'overwritePolicy': 'failIfExists'}, role='required', risk='write'),
)

EXCEL_FORMAT_VALIDATORS = {
    'absoluteXlsxPath': _xlsx_path, 'excelFormula': _formula,
    'excelLiteral': lambda value: all(c in {'\t', '\n'} or unicodedata.category(c) not in {'Cc', 'Cs'} for c in value),
}
EXCEL_FORMAT_VALIDATORS.update({
    'absolutePdfPath': lambda value: value.lower().endswith('.pdf') and _xlsx_path(value[:-4]+'.xlsx'),
})
EXCEL_TARGET_CONTRACT_SET = ApplicationContractSet(application='excel', contracts=_TARGET_CONTRACTS,
                                                   format_validators=EXCEL_FORMAT_VALIDATORS)
# Admitted after Windows WPS 12.0.0.28505 live acceptance on 2026-09-05.
# Evidence: src/test/resources/wps_skills/excel/type_library/EVIDENCE.md.
_EXCEL_PRODUCTION_ACTIONS = frozenset({
    'saveAs', 'createWorkbook', 'exportPdf',
    'openWorkbook', 'getWorkbookInfo', 'listWorksheets', 'readRange',
    'writeRange', 'setFormulas', 'calculateRange', 'formatRange', 'save',
    'getWorksheetInfo', 'addWorksheet', 'renameWorksheet', 'copyWorksheet', 'moveWorksheet',
    'deleteWorksheet', 'insertRows', 'deleteRows', 'insertColumns', 'deleteColumns',
    'clearRange', 'mergeRange', 'unmergeRange', 'sortRange', 'filterRange', 'clearFilter',
    'setRowHeight', 'setColumnWidth', 'autoFitColumns', 'replaceInRange', 'copyRange', 'findInRange',
})
EXCEL_PRODUCTION_CONTRACT_SET = ApplicationContractSet(
    application='excel',
    contracts=tuple(c for c in _TARGET_CONTRACTS if c.name in _EXCEL_PRODUCTION_ACTIONS),
    format_validators=EXCEL_FORMAT_VALIDATORS,
)
EXCEL_PRODUCTION_ACTION_INDEX = EXCEL_PRODUCTION_CONTRACT_SET.action_index()
