"""Formal Task CLI with real Word resources and a document-free adapter."""

from pathlib import Path
import sys
import os
import time
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[5]
sys.path[:0] = [str(ROOT / "main/python"), str(ROOT / "test/python")]

from wps_skills.cli.call import main
from wps_skills.client import task_client
from wps_skills.core.action_runtime import (AcquiredDocument, ActionContract, ApplicationContractSet,
                                           ControllerResult, PreparedDocumentAcquisition)
from wps_skills.word.task.executor import WordTask
from tests.core.resource_fixture import FakeDocumentCoordinator


events, mode = Path(sys.argv[1]), sys.argv[2]


class Adapter:
    application = "word"

    def prepare_establish(self, command):
        return PreparedDocumentAcquisition(coordination_identity="fixture", application_state=None)

    def establish(self, prepared, command):
        self.record(command)
        return AcquiredDocument(document=object(), data={"echo": dict(command.params)})

    def record(self, command):
        with events.open("a", encoding="utf-8") as stream:
            stream.write(command.address.action + "\n")

    def is_live(self, document):
        return True

    def handle(self, document, command):
        self.record(command)
        if mode == "owner_loss":
            os._exit(9)
        return ControllerResult.succeeded(data={"echo": dict(command.params)}, controller_state="usable",
                                           binding_disposition="unchanged")

    def handle_none(self, command):
        raise AssertionError("no none Actions")

    def close(self):
        if mode == "cleanup_loss":
            os._exit(9)
        events.with_suffix(".closed").write_text("closed")
        return True


# Only fixture response schemas differ; plan admission still uses production names.
from wps_skills.word.contracts import WORD_PRODUCTION_CONTRACT_SET
from dataclasses import replace
contracts = ApplicationContractSet(application="word", format_validators=WORD_PRODUCTION_CONTRACT_SET._format_validators, contracts=[
    replace(c, semantic_validator=None, result={"type": "object", "properties": {"echo": {"type": "object"}}, "required": ["echo"]})
    for c in WORD_PRODUCTION_CONTRACT_SET.contracts
])


def factory(**kwargs):
    return WordTask(**kwargs, adapter=Adapter(), coordinator=FakeDocumentCoordinator())


with patch("wps_skills.word.windows.task_factory.build_task", factory), patch.object(task_client, "contracts_for", return_value=contracts):
    code = main(sys.argv[3:])
    assert not any(name in sys.modules for name in ("wps_skills.core.action_session", "wps_skills.host.session_host", "wps_skills.client.session_client"))
    raise SystemExit(code)
