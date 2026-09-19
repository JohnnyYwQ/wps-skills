import unittest
import threading
import time

from wps_skills.core.action_runtime import (
    AcquiredDocument,
    ActionAddress,
    ActionContract,
    ActionError,
    ActionRequest,
    _RuntimeDisposition,
    ExecutionResources,
    _RuntimeTurn,
    ApplicationAdapter,
    ApplicationContractSet,
    ControllerResult,
    ControllerCommand,
    DocumentResourceCleanup,
    DefiniteEstablishFailure,
    PreparedDocumentAcquisition,
    ProcessCleanup,
    RequiredBindingFailure,
    TraceContext,
    UnprovableEstablishFailure,
)


from tests.core.resource_fixture import FakeApplicationAdapter, FakeDocumentCoordinator


def excel_contracts(*contracts):
    return ApplicationContractSet(
        application="excel",
        contracts=contracts,
        allow_incomplete=True,
    )


class ExecutionResourcesBindingTests(unittest.TestCase):
    def test_execution_resources_owns_and_reports_launcher_cleanup_once(self):
        class Launcher:
            def __init__(self):
                self.close_count = 0

            def close(self):
                self.close_count += 1
                return (ProcessCleanup(
                    pid=42,
                    cleanup_steps=("already_exited",),
                    released=True,
                ),)

        launcher = Launcher()
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=FakeApplicationAdapter(),
            coordinator=FakeDocumentCoordinator(),

            launcher=launcher,
        )

        first = resources.close()
        second = resources.close()

        self.assertIs(first, second)
        self.assertEqual(1, launcher.close_count)
        self.assertEqual(42, first.cleanup.processes[0].pid)

    def test_closed_core_types_reject_unknown_tags(self):
        with self.assertRaises(ValueError):
            ActionAddress(app="pages", action="inspectDocument")
        with self.assertRaises(ValueError):
            ControllerResult(
                outcome="maybe",
                controller_state="usable",
                binding_disposition="unchanged",
                error=ActionError(code="BAD", message="bad tag"),
            )
        with self.assertRaises(ValueError):
            ControllerResult(
                outcome="failed",
                controller_state="usable",
                binding_disposition="unchanged",
                error="bad",
            )
        response = _RuntimeDisposition(
            outcome="failed",
            address=ActionAddress(app="excel", action="inspectDocument"),

            trace=TraceContext(trace_id="trace-1", trace_log=None),
            error=ActionError(code="FAILED", message="failed"),
        )
        with self.assertRaises(ValueError):
            _RuntimeTurn(disposition=response, continuation="retry")
        with self.assertRaises(ValueError):
            DocumentResourceCleanup(state="maybe_released")

    def test_contract_set_rejects_duplicate_action_names(self):
        contract = ActionContract(
            name="inspectDocument",
            binding_role="required",
            risk="read",
            parameters={"type": "object"},
            result={"type": "object"},
        )
        with self.assertRaisesRegex(ValueError, "duplicate Action"):
            excel_contracts(contract, contract)

    def test_contract_set_rejects_invalid_binding_role_and_risk(self):
        invalid_contracts = (
            ActionContract(
                name="optionalDocument",
                binding_role="optional",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
            ),
            ActionContract(
                name="unsafeRisk",
                binding_role="required",
                risk="execute",
                parameters={"type": "object"},
                result={"type": "object"},
            ),
        )
        for contract in invalid_contracts:
            with self.subTest(contract=contract.name):
                with self.assertRaises(ValueError):
                    excel_contracts(contract)

    def test_resources_construction_rejects_mismatched_contracts_or_adapter(self):
        contract = ActionContract(
            name="inspectDocument",
            binding_role="required",
            risk="read",
            parameters={"type": "object"},
            result={"type": "object"},
        )
        with self.assertRaisesRegex(ValueError, "Contract Set application"):
            ExecutionResources(
                application="excel",
                contracts=ApplicationContractSet(
                    application="ppt",
                    contracts=(contract,),
                    allow_incomplete=True,
                ),
                adapter=FakeApplicationAdapter(),
                coordinator=FakeDocumentCoordinator(),

            )
        adapter = FakeApplicationAdapter()
        adapter.application = "ppt"
        with self.assertRaisesRegex(ValueError, "Adapter application"):
            ExecutionResources(
                application="excel",
                contracts=excel_contracts(contract),
                adapter=adapter,
                coordinator=FakeDocumentCoordinator(),

            )

        class IncompleteAdapter:
            application = "excel"

        self.assertNotIsInstance(IncompleteAdapter(), ApplicationAdapter)
        with self.assertRaisesRegex(ValueError, "Application Adapter seam"):
            ExecutionResources(
                application="excel",
                contracts=excel_contracts(contract),
                adapter=IncompleteAdapter(),
                coordinator=FakeDocumentCoordinator(),

            )

    def test_fresh_resources_rejects_required_action_without_dispatch(self):
        adapter = FakeApplicationAdapter()
        coordinator = FakeDocumentCoordinator()
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="inspectDocument",
                binding_role="required",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,

        )

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        self.assertEqual("failed", turn.disposition.outcome)
        self.assertEqual(
            "TASK_DOCUMENT_NOT_BOUND",
            turn.disposition.error.code,
        )
        self.assertEqual("continue", turn.continuation)
        self.assertEqual([], adapter.calls)
        self.assertEqual([], coordinator.calls)

    def test_resources_rejects_different_application_without_dispatch(self):
        adapter = FakeApplicationAdapter()
        coordinator = FakeDocumentCoordinator()
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="inspectDocument",
                binding_role="required",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,

        )

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="ppt", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        self.assertEqual("failed", turn.disposition.outcome)
        self.assertEqual("TASK_APP_MISMATCH", turn.disposition.error.code)
        self.assertEqual("continue", turn.continuation)
        self.assertEqual([], adapter.calls)
        self.assertEqual([], coordinator.calls)

    def test_unknown_local_action_is_a_nonterminal_action_failure(self):
        adapter = FakeApplicationAdapter()
        coordinator = FakeDocumentCoordinator()
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(),
            adapter=adapter,
            coordinator=coordinator,

        )

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="notAnAction"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        self.assertEqual("failed", turn.disposition.outcome)
        self.assertEqual("UNKNOWN_ACTION", turn.disposition.error.code)
        self.assertEqual("continue", turn.continuation)
        self.assertEqual([], adapter.calls)
        self.assertEqual([], coordinator.calls)

    def test_none_action_executes_without_document_resources_and_preserves_unbound_state(self):
        adapter = FakeApplicationAdapter()
        coordinator = FakeDocumentCoordinator()
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(
                ActionContract(
                    name="status",
                    binding_role="none",
                    risk="read",
                    parameters={"type": "object"},
                    result={
                        "type": "object",
                        "properties": {"status": {"type": "string"}},
                        "required": ["status"],
                        "additionalProperties": False,
                    },
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=coordinator,

        )

        status = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="status"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        required = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("succeeded", status.disposition.outcome)
        self.assertEqual({"status": "ready"}, status.disposition.data)
        self.assertEqual("continue", status.continuation)
        self.assertEqual("TASK_DOCUMENT_NOT_BOUND", required.disposition.error.code)
        self.assertEqual([], coordinator.calls)
        self.assertEqual(
            1,
            len([call for call in adapter.calls if call[0] == "handle_none"]),
        )

    def test_successful_establish_commits_document_and_lease_before_response(self):
        events = []
        document = object()
        adapter = FakeApplicationAdapter(document=document, events=events)
        coordinator = FakeDocumentCoordinator(events=events)
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,

        )

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        self.assertEqual("succeeded", turn.disposition.outcome)
        self.assertEqual({"created": True}, turn.disposition.data)
        self.assertEqual("continue", turn.continuation)
        self.assertEqual(
            [
                "adapter.prepare_establish",
                "coordinator.begin",
                "adapter.establish",
                "coordinator.commit",
            ],
            events,
        )
        self.assertEqual(
            [("commit", coordinator.guard, document)],
            [call for call in coordinator.calls if call[0] == "commit"],
        )

    def test_establishing_state_is_recorded_before_adapter_acquisition(self):
        events = []
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=FakeApplicationAdapter(document=object(), events=events),
            coordinator=FakeDocumentCoordinator(events=events),

        )

        resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(
                trace_id="trace-1",
                trace_log=None,
                event_sink=events.append,
            ),
        )

        self.assertLess(
            events.index("binding.establishing"),
            events.index("adapter.prepare_establish"),
        )

    def test_required_request_rejects_all_document_routing_decoys(self):
        document = object()
        adapter = FakeApplicationAdapter(document=document)
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={
                        "type": "object",
                        "properties": {},
                        "additionalProperties": False,
                    },
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),

        )
        resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        decoys = {
            "path": "C:/wrong.docx",
            "displayName": "wrong.docx",
            "activeDocument": True,
            "candidate": {"id": "candidate-1"},
            "registryKey": "registry-1",
            "documentId": "document-1",
        }
        for index, (name, value) in enumerate(decoys.items(), start=2):
            with self.subTest(decoy=name):
                turn = resources.execute(
                    ActionRequest(
                        address=ActionAddress(
                            app="excel",
                            action="inspectDocument",
                        ),
                        params={name: value},
                    ),
                    TraceContext(
                        trace_id=f"trace-{index}",
                        trace_log=None,
                    ),
                )
                self.assertEqual("failed", turn.disposition.outcome)
                self.assertEqual("INVALID_PARAMS", turn.disposition.error.code)
                self.assertEqual("continue", turn.continuation)
        self.assertEqual(
            [],
            [call for call in adapter.calls if call[0] == "handle"],
        )
        self.assertEqual(
            [],
            [call for call in adapter.calls if call[0] == "is_live"],
        )

    def test_required_action_rejects_illegal_controller_binding_matrix(self):
        document = object()
        adapter = FakeApplicationAdapter(
            document=document,
            handler_script=[ControllerResult.succeeded(
                data={"text": "wrong transition"},
                controller_state="usable",
                binding_disposition="established",
            )],
        )
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),

        )
        resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("failed", turn.disposition.outcome)
        self.assertEqual("INVALID_RESULT", turn.disposition.error.code)
        self.assertEqual("terminate", turn.continuation)

    def test_read_only_invalid_success_data_maps_to_failed_without_losing_binding(self):
        document = object()
        adapter = FakeApplicationAdapter(
            document=document,
            handler_script=[
                ControllerResult.succeeded(
                    data={},
                    controller_state="usable",
                    binding_disposition="unchanged",
                ),
                ControllerResult.succeeded(
                    data={"text": "still bound"},
                    controller_state="usable",
                    binding_disposition="unchanged",
                ),
            ],
        )
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={
                        "type": "object",
                        "properties": {"text": {"type": "string"}},
                        "required": ["text"],
                        "additionalProperties": False,
                    },
                ),
            ),
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),

        )
        resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        request = ActionRequest(
            address=ActionAddress(app="excel", action="inspectDocument"),
            params={},
        )

        invalid = resources.execute(
            request,
            TraceContext(trace_id="trace-2", trace_log=None),
        )
        valid = resources.execute(
            request,
            TraceContext(trace_id="trace-3", trace_log=None),
        )

        self.assertEqual("failed", invalid.disposition.outcome)
        self.assertEqual("INVALID_RESULT", invalid.disposition.error.code)
        self.assertEqual("continue", invalid.continuation)
        self.assertEqual("succeeded", valid.disposition.outcome)
        self.assertEqual({"text": "still bound"}, valid.disposition.data)

    def test_invalid_establish_data_is_unknown_terminal_after_binding_commit(self):
        document = object()
        adapter = FakeApplicationAdapter(
            document=document,
            establish_script=[AcquiredDocument(document=document, data={})],
        )
        coordinator = FakeDocumentCoordinator()
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={
                    "type": "object",
                    "properties": {"created": {"type": "boolean"}},
                    "required": ["created"],
                    "additionalProperties": False,
                },
            )),
            adapter=adapter,
            coordinator=coordinator,

        )

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        cleanup = resources.close()

        self.assertEqual("unknown", turn.disposition.outcome)
        self.assertEqual("INVALID_RESULT", turn.disposition.error.code)
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual("succeeded", cleanup.outcome)
        self.assertEqual(
            [("release", None, document, coordinator.lease)],
            [call for call in coordinator.calls if call[0] == "release"],
        )

    def test_required_handler_receives_bound_document_after_active_document_changes(self):
        bound_document = object()
        other_document = object()
        adapter = FakeApplicationAdapter(document=bound_document)
        coordinator = FakeDocumentCoordinator()
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=coordinator,

        )
        resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        adapter.active_document = other_document
        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("succeeded", turn.disposition.outcome)
        self.assertEqual({"text": "bound document"}, turn.disposition.data)
        self.assertEqual("continue", turn.continuation)
        self.assertEqual(
            [bound_document],
            [call[1] for call in adapter.calls if call[0] == "handle"],
        )

    def test_adapter_receives_one_immutable_controller_command_per_action(self):
        document = object()
        adapter = FakeApplicationAdapter(document=document)
        request_ids = iter(["request-1", "request-2"])
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),

            request_id_factory=lambda: next(request_ids),
            clock=lambda: 100.0,
            action_timeout_seconds=120,
        )
        resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        required_params = {"nested": {"ids": [1]}}
        resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="inspectDocument"),
                params=required_params,
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )
        required_params["nested"]["ids"].append(2)

        establish_command = next(
            call[1] for call in adapter.calls if call[0] == "establish"
        )
        required_document, required_command = next(
            (call[1], call[2])
            for call in adapter.calls
            if call[0] == "handle"
        )
        self.assertIsInstance(establish_command, ControllerCommand)
        self.assertEqual(
            ActionAddress(app="excel", action="createDocument"),
            establish_command.address,
        )
        self.assertEqual("request-1", establish_command.context.request_id)
        self.assertEqual("trace-1", establish_command.context.trace_id)
        self.assertEqual(220.0, establish_command.context.deadline_at)
        self.assertIs(document, required_document)
        self.assertEqual("request-2", required_command.context.request_id)
        self.assertEqual("trace-2", required_command.context.trace_id)
        self.assertEqual(220.0, required_command.context.deadline_at)
        self.assertEqual([1], list(required_command.params["nested"]["ids"]))
        with self.assertRaises(TypeError):
            required_command.params["nested"]["new"] = True

    def test_bound_resources_rejects_second_establish_without_dispatch(self):
        adapter = FakeApplicationAdapter(document=object())
        coordinator = FakeDocumentCoordinator()
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,

        )
        request = ActionRequest(
            address=ActionAddress(app="excel", action="createDocument"),
            params={},
        )
        resources.execute(
            request,
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        turn = resources.execute(
            request,
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("failed", turn.disposition.outcome)
        self.assertEqual(
            "TASK_DOCUMENT_ALREADY_BOUND",
            turn.disposition.error.code,
        )
        self.assertEqual("continue", turn.continuation)
        establish_calls = [
            call for call in adapter.calls if call[0] == "establish"
        ]
        self.assertEqual(1, len(establish_calls))
        self.assertEqual(
            "createDocument",
            establish_calls[0][1].address.action,
        )
        self.assertEqual(
            1,
            len([call for call in coordinator.calls if call[0] == "begin"]),
        )

    def test_definite_establish_failure_releases_resources_and_allows_retry(self):
        document = object()
        adapter = FakeApplicationAdapter(
            document=document,
            establish_script=[
                DefiniteEstablishFailure(
                    code="DOCUMENT_NOT_FOUND",
                    message="The requested document was not found",
                ),
                AcquiredDocument(
                    document=document,
                    data={"created": True},
                ),
            ],
        )
        coordinator = FakeDocumentCoordinator()
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,

        )
        request = ActionRequest(
            address=ActionAddress(app="excel", action="createDocument"),
            params={},
        )

        failed = resources.execute(
            request,
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        retried = resources.execute(
            request,
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("failed", failed.disposition.outcome)
        self.assertEqual("DOCUMENT_NOT_FOUND", failed.disposition.error.code)
        self.assertEqual("continue", failed.continuation)
        self.assertEqual("succeeded", retried.disposition.outcome)
        self.assertEqual("continue", retried.continuation)
        self.assertEqual(
            [("release", coordinator.guard, None, None)],
            [call for call in coordinator.calls if call[0] == "release"],
        )
        self.assertEqual(
            2,
            len([call for call in coordinator.calls if call[0] == "begin"]),
        )

    def test_definite_failure_with_quarantined_guard_is_terminal(self):
        coordinator = FakeDocumentCoordinator(release_state="quarantined")
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=FakeApplicationAdapter(establish_script=[
                DefiniteEstablishFailure(
                    code="DOCUMENT_NOT_FOUND",
                    message="The requested document was not found",
                ),
            ]),
            coordinator=coordinator,

        )

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        cleanup = resources.close()

        self.assertEqual("failed", turn.disposition.outcome)
        self.assertEqual(
            "DOCUMENT_BINDING_UNAVAILABLE",
            turn.disposition.error.code,
        )
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual("failed", cleanup.outcome)
        self.assertEqual("quarantined", cleanup.cleanup.document_resources.state)
        self.assertEqual(
            1,
            len([call for call in coordinator.calls if call[0] == "release"]),
        )

    def test_close_does_not_repeat_guard_release_already_in_progress(self):
        release_entered = threading.Event()
        resume_release = threading.Event()
        coordinator = FakeDocumentCoordinator(
            release_wait=resume_release,
            release_entered=release_entered,
        )
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=FakeApplicationAdapter(establish_script=[
                DefiniteEstablishFailure(
                    code="DOCUMENT_NOT_FOUND",
                    message="The requested document was not found",
                ),
            ]),
            coordinator=coordinator,

        )
        turns = []
        worker = threading.Thread(target=lambda: turns.append(resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )))
        worker.start()
        self.assertTrue(release_entered.wait(timeout=1))

        cleanup = resources.close()
        resume_release.set()
        worker.join(timeout=2)

        self.assertFalse(worker.is_alive())
        self.assertEqual("failed", cleanup.outcome)
        self.assertEqual(
            "release_unconfirmed",
            cleanup.cleanup.document_resources.state,
        )
        self.assertEqual("terminate", turns[0].continuation)
        self.assertEqual(
            1,
            len([call for call in coordinator.calls if call[0] == "release"]),
        )

    def test_definite_guard_acquisition_failure_remains_unbound_without_release(self):
        coordinator = FakeDocumentCoordinator(
            begin_error=DefiniteEstablishFailure(
                code="DOCUMENT_LEASE_CONFLICT",
                message="Another Session owns this document",
            ),
        )
        adapter = FakeApplicationAdapter(document=object())
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,

        )

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        cleanup = resources.close()

        self.assertEqual("failed", turn.disposition.outcome)
        self.assertEqual("DOCUMENT_LEASE_CONFLICT", turn.disposition.error.code)
        self.assertEqual("continue", turn.continuation)
        self.assertEqual("not_acquired", cleanup.cleanup.document_resources.state)
        self.assertEqual([], adapter.calls)
        self.assertEqual(
            [],
            [call for call in coordinator.calls if call[0] == "release"],
        )

    def test_unknown_unprovable_establish_is_terminal_and_cleanup_releases_partial_guard(self):
        partial_document = object()
        adapter = FakeApplicationAdapter(
            establish_script=[UnprovableEstablishFailure(
                outcome="unknown",
                code="RESPONSE_LOST",
                message="No reliable establish response was received",
                partial_document=partial_document,
            )],
        )
        coordinator = FakeDocumentCoordinator()
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,

        )

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        first_cleanup = resources.close()
        second_cleanup = resources.close()

        self.assertEqual("unknown", turn.disposition.outcome)
        self.assertEqual("RESPONSE_LOST", turn.disposition.error.code)
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual("succeeded", first_cleanup.outcome)
        self.assertEqual("released", first_cleanup.cleanup.document_resources.state)
        self.assertIs(first_cleanup, second_cleanup)
        self.assertEqual(
            [("release", coordinator.guard, partial_document, None)],
            [call for call in coordinator.calls if call[0] == "release"],
        )

    def test_failed_unprovable_establish_normalizes_binding_error_and_terminates(self):
        partial_document = object()
        adapter = FakeApplicationAdapter(
            establish_script=[UnprovableEstablishFailure(
                outcome="failed",
                code="OPEN_FAILED",
                message="The document may have opened without a provable binding",
                partial_document=partial_document,
            )],
        )
        coordinator = FakeDocumentCoordinator()
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,

        )

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        cleanup = resources.close()

        self.assertEqual("failed", turn.disposition.outcome)
        self.assertEqual(
            "DOCUMENT_BINDING_UNAVAILABLE",
            turn.disposition.error.code,
        )
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual("succeeded", cleanup.outcome)
        self.assertEqual(
            [("release", coordinator.guard, partial_document, None)],
            [call for call in coordinator.calls if call[0] == "release"],
        )

    def test_lease_commit_failure_is_terminal_and_cleanup_reclaims_partial_document(self):
        document = object()
        coordinator = FakeDocumentCoordinator(
            commit_error=RuntimeError("lease commit failed"),
        )
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=FakeApplicationAdapter(document=document),
            coordinator=coordinator,

        )

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        cleanup = resources.close()

        self.assertEqual("failed", turn.disposition.outcome)
        self.assertEqual(
            "DOCUMENT_BINDING_UNAVAILABLE",
            turn.disposition.error.code,
        )
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual("succeeded", cleanup.outcome)
        self.assertEqual(
            [("release", coordinator.guard, document, None)],
            [call for call in coordinator.calls if call[0] == "release"],
        )

    def test_required_binding_loss_preserves_failure_and_terminates(self):
        document = object()
        adapter = FakeApplicationAdapter(
            document=document,
            handler_script=[RequiredBindingFailure(
                disposition="lost",
                outcome="failed",
                code="DOCUMENT_CLOSED",
                message="The bound document is closed",
            )],
        )
        coordinator = FakeDocumentCoordinator()
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=coordinator,

        )
        resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )
        cleanup = resources.close()

        self.assertEqual("failed", turn.disposition.outcome)
        self.assertEqual("DOCUMENT_CLOSED", turn.disposition.error.code)
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual("succeeded", cleanup.outcome)
        self.assertEqual(
            [("release", None, document, coordinator.lease)],
            [call for call in coordinator.calls if call[0] == "release"],
        )

    def test_read_only_required_unprovable_unknown_maps_to_failed_and_terminates(self):
        document = object()
        adapter = FakeApplicationAdapter(
            document=document,
            handler_script=[RequiredBindingFailure(
                disposition="unprovable",
                outcome="unknown",
                code="RESPONSE_LOST",
                message="The read response could not be recovered",
            )],
        )
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),

        )
        resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("failed", turn.disposition.outcome)
        self.assertEqual("RESPONSE_LOST", turn.disposition.error.code)
        self.assertEqual("terminate", turn.continuation)

    def test_closed_bound_document_prevents_required_handler_and_terminates(self):
        document = object()
        adapter = FakeApplicationAdapter(document=document)
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),

        )
        resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        adapter.live = False

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("failed", turn.disposition.outcome)
        self.assertEqual("DOCUMENT_CLOSED", turn.disposition.error.code)
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual(
            [],
            [call for call in adapter.calls if call[0] == "handle"],
        )
        self.assertEqual(
            [("is_live", document)],
            [call for call in adapter.calls if call[0] == "is_live"],
        )

    def test_bound_cleanup_reports_quarantine_failure_and_releases_once(self):
        document = object()
        adapter = FakeApplicationAdapter(document=document)
        coordinator = FakeDocumentCoordinator(release_state="quarantined")
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,

        )
        resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        first = resources.close()
        second = resources.close()

        self.assertEqual("failed", first.outcome)
        self.assertEqual("TASK_CLEANUP_INCOMPLETE", first.error.code)
        self.assertEqual("quarantined", first.cleanup.document_resources.state)
        self.assertIs(first, second)
        self.assertEqual(
            [("release", None, document, coordinator.lease)],
            [call for call in coordinator.calls if call[0] == "release"],
        )

    def test_cleanup_release_exception_is_reported_and_not_retried(self):
        document = object()
        adapter = FakeApplicationAdapter(document=document)
        coordinator = FakeDocumentCoordinator(
            release_error=RuntimeError("coordinator unavailable"),
        )
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,

        )
        resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        first = resources.close()
        second = resources.close()

        self.assertEqual("failed", first.outcome)
        self.assertEqual(
            "release_unconfirmed",
            first.cleanup.document_resources.state,
        )
        self.assertEqual("TASK_CLEANUP_INCOMPLETE", first.error.code)
        self.assertIs(first, second)
        self.assertEqual(
            1,
            len([call for call in coordinator.calls if call[0] == "release"]),
        )

    def test_close_during_establish_is_bounded_and_quarantines_in_flight_work(self):
        entered = threading.Event()
        resume = threading.Event()
        document = object()

        class BlockingAdapter(FakeApplicationAdapter):
            def establish(self, prepared, command):
                self.calls.append(("establish", command))
                entered.set()
                resume.wait(timeout=2)
                return AcquiredDocument(
                    document=document,
                    data={"created": True},
                )

        coordinator = FakeDocumentCoordinator(release_state="quarantined")
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=BlockingAdapter(document=document),
            coordinator=coordinator,

        )
        turns = []
        worker = threading.Thread(target=lambda: turns.append(resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )))
        worker.start()
        self.assertTrue(entered.wait(timeout=1))

        started = time.monotonic()
        cleanup = resources.close()
        elapsed = time.monotonic() - started
        resume.set()
        worker.join(timeout=2)

        self.assertLess(elapsed, 0.5)
        self.assertFalse(worker.is_alive())
        self.assertEqual("failed", cleanup.outcome)
        self.assertEqual("quarantined", cleanup.cleanup.document_resources.state)
        self.assertEqual([True], coordinator.release_in_flight)
        self.assertEqual("terminate", turns[0].continuation)
        self.assertIs(cleanup, resources.close())
        self.assertEqual(
            1,
            len([call for call in coordinator.calls if call[0] == "release"]),
        )

    def test_close_while_guard_acquisition_blocks_retains_in_flight_fence(self):
        begin_entered = threading.Event()
        resume_begin = threading.Event()
        coordinator = FakeDocumentCoordinator(
            release_state="quarantined",
            begin_wait=resume_begin,
            begin_entered=begin_entered,
        )
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=FakeApplicationAdapter(document=object()),
            coordinator=coordinator,

        )
        turns = []
        worker = threading.Thread(target=lambda: turns.append(resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )))
        worker.start()
        self.assertTrue(begin_entered.wait(timeout=1))

        cleanup = resources.close()
        resume_begin.set()
        worker.join(timeout=2)

        self.assertFalse(worker.is_alive())
        self.assertEqual("failed", cleanup.outcome)
        self.assertEqual("quarantined", cleanup.cleanup.document_resources.state)
        self.assertEqual([True], coordinator.release_in_flight)
        self.assertEqual(
            [("release", None, None, None)],
            [call for call in coordinator.calls if call[0] == "release"],
        )
        self.assertEqual("terminate", turns[0].continuation)

    def test_cleanup_timeout_reports_unconfirmed_release_without_blocking(self):
        release = threading.Event()
        coordinator = FakeDocumentCoordinator(release_wait=release)
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=FakeApplicationAdapter(document=object()),
            coordinator=coordinator,

            cleanup_timeout_seconds=0.02,
        )
        resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        started = time.monotonic()
        cleanup = resources.close()
        elapsed = time.monotonic() - started
        release.set()

        self.assertLess(elapsed, 0.5)
        self.assertEqual("failed", cleanup.outcome)
        self.assertEqual(
            "release_unconfirmed",
            cleanup.cleanup.document_resources.state,
        )

    def test_unbound_cleanup_is_idempotent_and_stops_new_dispatch(self):
        adapter = FakeApplicationAdapter()
        coordinator = FakeDocumentCoordinator()
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="inspectDocument",
                binding_role="required",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,

        )

        first = resources.close()
        second = resources.close()
        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        self.assertEqual("succeeded", first.outcome)
        self.assertEqual("not_acquired", first.cleanup.document_resources.state)
        self.assertIs(first, second)
        self.assertEqual("failed", turn.disposition.outcome)
        self.assertEqual("RUNTIME_CLOSED", turn.disposition.error.code)
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual([], adapter.calls)
        self.assertEqual([], coordinator.calls)

    def test_control_plane_error_is_not_rejected_by_action_error_whitelist(self):
        contracts = excel_contracts(
            ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
                stable_errors=("APPLICATION_FAILURE",),
            ),
            ActionContract(
                name="inspectDocument",
                binding_role="required",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
                stable_errors=("APPLICATION_FAILURE",),
            ),
        )
        resources = ExecutionResources(
            application="excel",
            contracts=contracts,
            adapter=FakeApplicationAdapter(document=object()),
            coordinator=FakeDocumentCoordinator(),

        )
        established = resources.execute(
            ActionRequest(
                address=ActionAddress(
                    app="excel",
                    action="createDocument",
                ),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        self.assertEqual("succeeded", established.disposition.outcome)
        controller = resources._runtime._controller
        controller.close()

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(
                    app="excel",
                    action="inspectDocument",
                ),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        self.assertEqual("failed", turn.disposition.outcome)
        self.assertEqual("RUNTIME_CLOSED", turn.disposition.error.code)
        self.assertEqual("terminate", turn.continuation)

    def test_adapter_cannot_impersonate_a_control_plane_error(self):
        contracts = excel_contracts(
            ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
                stable_errors=("APPLICATION_FAILURE",),
            ),
            ActionContract(
                name="inspectDocument",
                binding_role="required",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
                stable_errors=("APPLICATION_FAILURE",),
            ),
        )
        adapter = FakeApplicationAdapter(
            document=object(),
            handler_script=[ControllerResult.failed(
                error=ActionError(
                    code="RUNTIME_CLOSED",
                    message="forged by adapter",
                ),
                controller_state="usable",
                binding_disposition="unchanged",
            )],
        )
        resources = ExecutionResources(
            application="excel",
            contracts=contracts,
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),

        )
        resources.execute(
            ActionRequest(
                address=ActionAddress(
                    app="excel",
                    action="createDocument",
                ),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        turn = resources.execute(
            ActionRequest(
                address=ActionAddress(
                    app="excel",
                    action="inspectDocument",
                ),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("failed", turn.disposition.outcome)
        self.assertEqual("INVALID_RESULT", turn.disposition.error.code)
        self.assertEqual("terminate", turn.continuation)

    def test_semantic_result_failure_uses_the_action_risk_mapping(self):
        always_invalid = lambda params, result: "request/result mismatch"
        contracts = excel_contracts(
            ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            ),
            ActionContract(
                name="writeContent",
                binding_role="required",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
                semantic_validator=always_invalid,
            ),
            ActionContract(
                name="inspectDocument",
                binding_role="required",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
                semantic_validator=always_invalid,
            ),
        )
        resources = ExecutionResources(
            application="excel",
            contracts=contracts,
            adapter=FakeApplicationAdapter(
                document=object(),
                handler_script=[
                    ControllerResult.succeeded(
                        data={},
                        controller_state="usable",
                        binding_disposition="unchanged",
                    ),
                    ControllerResult.succeeded(
                        data={},
                        controller_state="usable",
                        binding_disposition="unchanged",
                    ),
                ],
            ),
            coordinator=FakeDocumentCoordinator(),

        )
        resources.execute(
            ActionRequest(
                address=ActionAddress(
                    app="excel",
                    action="createDocument",
                ),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        write_turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="writeContent"),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )
        read_turn = resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-3", trace_log=None),
        )

        self.assertEqual("unknown", write_turn.disposition.outcome)
        self.assertEqual("INVALID_RESULT", write_turn.disposition.error.code)
        self.assertEqual("failed", read_turn.disposition.outcome)
        self.assertEqual("INVALID_RESULT", read_turn.disposition.error.code)

    def test_close_before_lazy_controller_creation_prevents_dispatch(self):
        request_id_entered = threading.Event()
        resume_request_id = threading.Event()

        def request_id_factory():
            request_id_entered.set()
            resume_request_id.wait(timeout=2)
            return "request-1"

        adapter = FakeApplicationAdapter()
        resources = ExecutionResources(
            application="excel",
            contracts=excel_contracts(ActionContract(
                name="status",
                binding_role="none",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),

            request_id_factory=request_id_factory,
        )
        turns = []
        worker = threading.Thread(target=lambda: turns.append(resources.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="status"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )))
        worker.start()
        self.assertTrue(request_id_entered.wait(timeout=1))

        cleanup = resources.close()
        resume_request_id.set()
        worker.join(timeout=2)

        self.assertEqual("not_acquired", cleanup.cleanup.document_resources.state)
        self.assertFalse(worker.is_alive())
        self.assertEqual("failed", turns[0].disposition.outcome)
        self.assertEqual("RUNTIME_CLOSED", turns[0].disposition.error.code)
        self.assertEqual("terminate", turns[0].continuation)
        self.assertEqual([], adapter.calls)


if __name__ == "__main__":
    unittest.main()
