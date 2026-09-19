"""Document-free adapters and coordinators for shared resource tests."""

from wps_skills.core.action_runtime import (
    PreparedDocumentAcquisition, AcquiredDocument, ControllerResult, DocumentResourceCleanup,
)

class FakeApplicationAdapter:
    application = "excel"

    def __init__(
        self,
        *,
        document=None,
        events=None,
        establish_script=None,
        handler_script=None,
    ):
        self.calls = []
        self.document = document
        self.events = events if events is not None else []
        self.active_document = document
        self.live = True
        self.establish_script = list(establish_script or [])
        self.handler_script = list(handler_script or [])

    def prepare_establish(self, command):
        self.events.append("adapter.prepare_establish")
        return PreparedDocumentAcquisition(
            coordination_identity="fake-document-identity",
            application_state=self,
        )

    def establish(self, prepared, command):
        if prepared.application_state is not self:
            raise ValueError("another acquisition plan")
        self.calls.append(("establish", command))
        self.events.append("adapter.establish")
        if self.establish_script:
            result = self.establish_script.pop(0)
            if isinstance(result, BaseException):
                raise result
            return result
        return AcquiredDocument(
            document=self.document,
            data={"created": True},
        )

    def is_live(self, document):
        self.calls.append(("is_live", document))
        return self.live

    def handle(self, document, command):
        self.calls.append(("handle", document, command))
        if self.handler_script:
            result = self.handler_script.pop(0)
            if isinstance(result, BaseException):
                raise result
            return result
        return ControllerResult.succeeded(
            data={"text": "bound document"},
            controller_state="usable",
            binding_disposition="unchanged",
        )

    def handle_none(self, command):
        self.calls.append(("handle_none", command))
        return ControllerResult.succeeded(
            data={"status": "ready"},
            controller_state="usable",
            binding_disposition="unchanged",
        )

    def close(self):
        return True


class FakeDocumentCoordinator:
    def __init__(
        self,
        *,
        events=None,
        release_state="released",
        release_error=None,
        commit_error=None,
        begin_error=None,
        release_wait=None,
        begin_wait=None,
        begin_entered=None,
        release_entered=None,
    ):
        self.calls = []
        self.events = events if events is not None else []
        self.guard = object()
        self.lease = object()
        self.release_state = release_state
        self.release_error = release_error
        self.commit_error = commit_error
        self.begin_error = begin_error
        self.release_wait = release_wait
        self.begin_wait = begin_wait
        self.begin_entered = begin_entered
        self.release_entered = release_entered
        self.release_in_flight = []

    def begin(self, coordination_identity, context):
        self.calls.append(("begin",))
        self.events.append("coordinator.begin")
        if self.begin_entered is not None:
            self.begin_entered.set()
        if self.begin_wait is not None:
            self.begin_wait.wait(timeout=2)
        if self.begin_error is not None:
            raise self.begin_error
        return self.guard

    def commit(self, guard, document, context):
        self.calls.append(("commit", guard, document))
        self.events.append("coordinator.commit")
        if self.commit_error is not None:
            raise self.commit_error
        return self.lease

    def release(self, guard, document, lease, *, in_flight=False):
        self.calls.append(("release", guard, document, lease))
        self.events.append("coordinator.release")
        self.release_in_flight.append(in_flight)
        if self.release_entered is not None:
            self.release_entered.set()
        if self.release_wait is not None:
            self.release_wait.wait(timeout=2)
        if self.release_error is not None:
            raise self.release_error
        return DocumentResourceCleanup(state=self.release_state)


