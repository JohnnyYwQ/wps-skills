"""Application-owned mappings; shared transport never routes public Actions."""

from types import MappingProxyType
from typing import Mapping

from wps_skills.core.action_session import ActionError, ControllerResult
from wps_skills.windows.bridge_types import BackendActionFailure

PRIVATE_OPERATIONS = MappingProxyType({
    'getWorkbookInfo': 'inspect_workbook',
    'listWorksheets': 'list_worksheets',
    'readRange': 'read_cell_rectangle',
    'writeRange': 'write_literal_rectangle',
    'setFormulas': 'write_formula_rectangle',
    'calculateRange': 'calculate_rectangle',
    'formatRange': 'format_rectangle',
    'saveAs': 'saveAs',
    'exportPdf': 'exportPdf',
    'save': 'save_existing_workbook',
    'getWorksheetInfo': 'getWorksheetInfo',
    'addWorksheet': 'addWorksheet',
    'renameWorksheet': 'renameWorksheet',
    'copyWorksheet': 'copyWorksheet',
    'moveWorksheet': 'moveWorksheet',
    'deleteWorksheet': 'deleteWorksheet',
    'insertRows': 'insertRows',
    'deleteRows': 'deleteRows',
    'insertColumns': 'insertColumns',
    'deleteColumns': 'deleteColumns',
    'clearRange': 'clearRange',
    'mergeRange': 'mergeRange',
    'unmergeRange': 'unmergeRange',
    'sortRange': 'sortRange',
    'filterRange': 'filterRange',
    'clearFilter': 'clearFilter',
    'setRowHeight': 'setRowHeight',
    'setColumnWidth': 'setColumnWidth',
    'autoFitColumns': 'autoFitColumns',
    'replaceInRange': 'replaceInRange',
    'copyRange': 'copyRange',
    'findInRange': 'findInRange',

})


def _handler(operation):
    def invoke(backend, document, params, context):
        try:
            data = backend.invoke(document, operation, params, context)
        except BackendActionFailure as exc:
            factory = ControllerResult.failed if exc.outcome == 'failed' else ControllerResult.unknown
            return factory(error=ActionError(code=exc.code, message=exc.message),
                           controller_state='usable' if exc.binding_disposition == 'unchanged' else 'broken',
                           binding_disposition=exc.binding_disposition)
        if not isinstance(data, Mapping):
            raise TypeError('Excel bridge result must be an object')
        return ControllerResult.succeeded(data=data, controller_state='usable', binding_disposition='unchanged')
    return invoke


EXCEL_HANDLERS = MappingProxyType({action: _handler(operation) for action, operation in PRIVATE_OPERATIONS.items()})
