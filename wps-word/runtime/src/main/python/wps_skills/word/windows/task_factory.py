"""Assemble one Windows Word Task with application-owned resources."""

from pathlib import Path

class TaskStartupError(RuntimeError):
    """Word resource construction failed before document acquisition."""

    def __init__(self, message, *, cleanup=None):
        super().__init__(message)
        self.cleanup = cleanup


BRIDGE_SCRIPT = (
    Path(__file__).resolve().parents[4]
    / "resources" / "wps_skills" / "word" / "windows" / "word_bridge.ps1"
)


def build_task(
    *,
    task_id,
    contracts=None,
    action_timeout_seconds=60,
):
    from wps_skills.word.task.executor import WordTask
    from wps_skills.windows.powershell_bridge import JsonLineBridgeTransport
    from wps_skills.windows.document_coordinator import WindowsDocumentCoordinator
    from wps_skills.windows.owned_process import WindowsOwnedProcessLauncher
    from wps_skills.word.windows.backend import WindowsWordBackend
    from wps_skills.windows.bridge_runtime import LazyWindowsBridge
    from wps_skills.word.actions.adapter import WordAdapter
    from wps_skills.word.contracts import WORD_PRODUCTION_CONTRACT_SET
    from wps_skills.word.actions.registry import WORD_HANDLERS

    contracts = WORD_PRODUCTION_CONTRACT_SET if contracts is None else contracts
    launcher = None
    try:
        launcher = WindowsOwnedProcessLauncher()
        bridge = LazyWindowsBridge(
            launcher=launcher,
            script_path=BRIDGE_SCRIPT,
            transport_factory=JsonLineBridgeTransport,
        )
        backend = WindowsWordBackend(bridge=bridge)
        adapter = WordAdapter(
            backend=backend,
            handlers=WORD_HANDLERS,
            contracts=contracts,
        )
        return WordTask(
            contracts=contracts,
            adapter=adapter,
            coordinator=WindowsDocumentCoordinator(bridge=bridge),
            task_id=task_id,
            launcher=launcher,
            action_timeout_seconds=action_timeout_seconds,
        )
    except (Exception, KeyboardInterrupt) as exc:
        cleanup = None
        if launcher is not None:
            try:
                processes = launcher.close()
                failed = any(not process.released for process in processes)
                cleanup = {"outcome": "failed" if failed else "succeeded",
                           "error": "Startup resources were not fully released" if failed else None}
            except (Exception, KeyboardInterrupt) as cleanup_error:
                cleanup = {"outcome": "unknown", "error": str(cleanup_error)}
        if isinstance(exc, KeyboardInterrupt):
            exc.cleanup = cleanup
            raise
        raise TaskStartupError(
            "the production Word Task could not be constructed", cleanup=cleanup,
        ) from exc
