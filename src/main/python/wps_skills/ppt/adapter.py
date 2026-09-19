"""Ppt Application Adapter; one Task binds one exact new or existing Presentation."""

from wps_skills.core.action_runtime import (
    AcquiredDocument, ControllerCommand, PreparedDocumentAcquisition,
)
from wps_skills.ppt.handlers import PPT_HANDLERS
from wps_skills.ppt.windows.backend import PptPreparation


class PptAdapter:
    application = 'ppt'

    def __init__(self, *, backend, contracts, handlers=PPT_HANDLERS):
        if contracts.application != self.application:
            raise ValueError('Ppt contracts must belong to ppt')
        required = {c.name for c in contracts.contracts if c.binding_role == 'required'}
        if not required.issubset(handlers) or any(not callable(handlers[name]) for name in required):
            raise ValueError('Ppt contracts require complete executable handlers')
        if any(c.binding_role == 'establish' and c.name not in {'openPresentation', 'createPresentation'} for c in contracts.contracts):
            raise ValueError('Ppt supports only explicit open or create establishment')
        self.contracts = contracts
        self._backend = backend
        self._handlers = {name: handlers[name] for name in required}

    @staticmethod
    def _command(command):
        if not isinstance(command, ControllerCommand) or command.address.app != 'ppt':
            raise ValueError('Ppt Adapter requires an Ppt Controller Command')

    def prepare_establish(self, command):
        self._command(command)
        if command.address.action not in {'openPresentation', 'createPresentation'}:
            raise ValueError('Action is not an Ppt establish Action')
        state = (self._backend.prepare_create(command.context) if command.address.action == 'createPresentation'
                 else self._backend.prepare_open(command.params['path'], command.context))
        if not isinstance(state, PptPreparation):
            raise TypeError('Ppt backend returned an invalid preparation')
        return PreparedDocumentAcquisition(coordination_identity=state.coordination_identity, application_state=state)

    def establish(self, prepared, command):
        self._command(command)
        if command.address.action not in {'openPresentation', 'createPresentation'} or not isinstance(prepared, PreparedDocumentAcquisition):
            raise ValueError('Ppt requires its prepared open or create acquisition')
        state = prepared.application_state
        if not isinstance(state, PptPreparation) or state.path != command.params.get('path') or state.coordination_identity != prepared.coordination_identity:
            raise ValueError('Ppt acquisition identity does not match')
        document, data = (self._backend.create(state, command.context) if command.address.action == 'createPresentation'
                          else self._backend.open(state, command.context))
        return AcquiredDocument(document=document, data=data)

    def is_live(self, document):
        return self._backend.is_live(document)

    def handle(self, document, command):
        self._command(command)
        return self._handlers[command.address.action](self._backend, document, command.params, command.context)

    def handle_none(self, command):
        self._command(command)
        raise ValueError('Ppt has no none Actions')

    def close(self):
        return self._backend.close()
