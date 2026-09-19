"""Assemble a Windows Task using an application-owned backend and adapter."""

from pathlib import Path

class TaskStartupError(RuntimeError):
    """Application resource construction failed before document acquisition."""

    def __init__(self, message, *, cleanup=None):
        super().__init__(message)
        self.cleanup = cleanup


def build_task(
    *,
    application,
    task_id,
    contracts=None,
    action_timeout_seconds=60,
):
    import importlib
    from wps_skills.core.task_executor import ApplicationTask
    from wps_skills.windows.powershell_bridge import JsonLineBridgeTransport
    from wps_skills.windows.document_coordinator import WindowsDocumentCoordinator
    from wps_skills.windows.owned_process import WindowsOwnedProcessLauncher
    from wps_skills.windows.bridge_runtime import LazyWindowsBridge
    from wps_skills.client.applications import contracts_for

    label = {"word": "Word", "excel": "Excel", "ppt": "Ppt"}[application]
    backend_type = getattr(importlib.import_module("wps_skills." + application + ".windows.backend"), "Windows" + label + "Backend")
    adapter_module = ".actions.adapter" if application == "word" else ".adapter"
    adapter_type = getattr(importlib.import_module("wps_skills." + application + adapter_module), label + "Adapter")
    contracts = contracts_for(application) if contracts is None else contracts
    script = Path(__file__).resolve().parents[3] / "resources" / "wps_skills" / application / "windows" / (application + "_bridge.ps1")
    launcher = None
    try:
        launcher = WindowsOwnedProcessLauncher()
        bridge = LazyWindowsBridge(
            launcher=launcher,
            script_path=script,
            transport_factory=JsonLineBridgeTransport,
        )
        backend = backend_type(bridge=bridge)
        adapter_options = {}
        if application == "word":
            from wps_skills.word.actions.registry import WORD_HANDLERS
            adapter_options["handlers"] = WORD_HANDLERS
        adapter = adapter_type(
            **adapter_options,
            backend=backend,
            contracts=contracts,
        )
        if application == "word":
            from wps_skills.word.task.executor import WordTask
            executor_type = WordTask
            extra = {}
        else:
            executor_type = ApplicationTask
            extra = {"application": application}
        return executor_type(
            **extra,
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
            "the production " + application + " Task could not be constructed", cleanup=cleanup,
        ) from exc
