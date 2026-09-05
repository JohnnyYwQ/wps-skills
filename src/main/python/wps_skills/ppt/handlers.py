"""Application-owned mappings; shared transport never routes public Actions."""

from types import MappingProxyType
from typing import Mapping

from wps_skills.core.action_session import ActionError, ControllerResult
from wps_skills.windows.bridge_types import BackendActionFailure

PRIVATE_OPERATIONS = MappingProxyType({name: name for name in (
    'getPresentationInfo',
    'listSlides',
    'getSlideInfo',
    'addSlide',
    'duplicateSlide',
    'moveSlide',
    'deleteSlide',
    'addTextBox',
    'addShape',
    'setShapeText',
    'formatText',
    'setShapeGeometry',
    'deleteShape',
    'save',
    'getShapeStyle',
    'formatShape',
    'formatParagraph',
    'setTextBoxLayout',
    'renameShape',
    'setShapeOrder',
    'alignShapes',
    'distributeShapes',
    'getSlideSettings',
    'setSlideSettings',
    'getSlideNotes',
    'setSlideNotes',
    'findText',
    'replaceText',
    'addImage',
    'addTable',
    'readTable',
    'writeTable',
)})


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
            raise TypeError('Ppt bridge result must be an object')
        return ControllerResult.succeeded(data=data, controller_state='usable', binding_disposition='unchanged')
    return invoke


PPT_HANDLERS = MappingProxyType({action: _handler(operation) for action, operation in PRIVATE_OPERATIONS.items()})
