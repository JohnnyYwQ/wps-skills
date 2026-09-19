"""Exact-presentation backend over the shared Windows bridge."""

from dataclasses import dataclass
from typing import Mapping, Optional

from wps_skills.core.action_runtime import DefiniteEstablishFailure, UnprovableEstablishFailure
from wps_skills.windows.bridge_types import BackendActionFailure, WindowsDocument


@dataclass(frozen=True)
class PptPreparation:
    preparation_id: str
    coordination_identity: str
    path: Optional[str]


def _fields(value, keys):
    if not isinstance(value, Mapping) or set(value) != set(keys):
        raise TypeError('Ppt bridge returned invalid fields')
    return value


def _text(value):
    if not isinstance(value, str) or not value:
        raise TypeError('Ppt bridge returned an invalid identifier')
    return value


def _establish_failure(exc):
    if exc.outcome == 'failed' and exc.binding_disposition == 'unchanged':
        raise DefiniteEstablishFailure(code=exc.code, message=exc.message) from exc
    raise UnprovableEstablishFailure(outcome=exc.outcome, code=exc.code, message=exc.message) from exc


class WindowsPptBackend:
    def __init__(self, *, bridge):
        self._bridge = bridge
        self._preparation = None
        self._document = None

    def prepare_open(self, path, context):
        if self._document is not None:
            raise ValueError('Ppt backend is already bound')
        try:
            result = _fields(self._bridge.execute('prepare_existing_document', {'path': path}, context),
                             {'preparationId', 'coordinationIdentity'})
        except BackendActionFailure as exc:
            _establish_failure(exc)
        self._preparation = PptPreparation(_text(result['preparationId']), _text(result['coordinationIdentity']), path)
        return self._preparation

    def prepare_create(self, context):
        if self._document is not None:
            raise ValueError('Ppt backend is already bound')
        try:
            result = _fields(self._bridge.execute('prepare_new_document', {}, context),
                             {'preparationId', 'coordinationIdentity'})
        except BackendActionFailure as exc:
            _establish_failure(exc)
        self._preparation = PptPreparation(_text(result['preparationId']), _text(result['coordinationIdentity']), None)
        return self._preparation

    def create(self, preparation, context):
        if preparation is not self._preparation or preparation.path is not None or self._document is not None:
            raise ValueError('Ppt backend requires its own new-document preparation')
        try:
            result = _fields(self._bridge.execute('acquire_new_document',
                             {'preparationId': preparation.preparation_id}, context),
                             {'documentId', 'documentState'})
        except BackendActionFailure as exc:
            _establish_failure(exc)
        self._document = WindowsDocument(_text(result['documentId']), None)
        return self._document, {'documentState': result['documentState']}

    def open(self, preparation, context):
        if preparation is not self._preparation or self._document is not None:
            raise ValueError('Ppt backend requires its own unused preparation')
        try:
            result = _fields(self._bridge.execute('acquire_existing_document',
                             {'preparationId': preparation.preparation_id}, context),
                             {'documentId', 'artifact', 'documentState'})
        except BackendActionFailure as exc:
            _establish_failure(exc)
        self._document = WindowsDocument(_text(result['documentId']), preparation.path)
        return self._document, {'artifact': result['artifact'], 'documentState': result['documentState']}

    def _bound(self, document):
        if document is not self._document:
            raise ValueError('Ppt backend received another presentation reference')

    def is_live(self, document):
        self._bound(document)
        result = _fields(self._bridge.execute('probe_bound_document',
                         {'documentId': document.bridge_document_id}, None), {'live'})
        if not isinstance(result['live'], bool):
            raise TypeError('Ppt liveness must be Boolean')
        return result['live']

    def invoke(self, document, operation, params, context):
        self._bound(document)
        return self._bridge.execute(operation, {'documentId': document.bridge_document_id,
                                    'operationArguments': dict(params)}, context)

    def close(self):
        return self._bridge.close()
