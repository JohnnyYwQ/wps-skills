"""Assemble one Windows Excel Action Session with application-owned resources."""

from pathlib import Path

from wps_skills.host.session_host import SessionStartupError


BRIDGE_SCRIPT = (
    Path(__file__).resolve().parents[4]
    / "resources" / "wps_skills" / "excel" / "windows" / "excel_bridge.ps1"
)


def build_session(*, session_id, contracts=None):
    from wps_skills.core.action_session import ActionSession
    from wps_skills.excel.adapter import ExcelAdapter
    from wps_skills.excel.contracts import EXCEL_PRODUCTION_CONTRACT_SET
    from wps_skills.windows.bridge_runtime import LazyWindowsBridge
    from wps_skills.windows.document_coordinator import WindowsDocumentCoordinator
    from wps_skills.excel.windows.backend import WindowsExcelBackend
    from wps_skills.windows.owned_process import WindowsOwnedProcessLauncher

    contracts = EXCEL_PRODUCTION_CONTRACT_SET if contracts is None else contracts
    if contracts is None:
        raise SessionStartupError("Excel implementation awaits live-WPS admission; no production Excel Application Contract Set is installed")
    launcher = None
    try:
        launcher = WindowsOwnedProcessLauncher()
        bridge = LazyWindowsBridge(launcher=launcher, script_path=BRIDGE_SCRIPT)
        return ActionSession(application="excel", contracts=contracts,
                             adapter=ExcelAdapter(backend=WindowsExcelBackend(bridge=bridge), contracts=contracts),
                             coordinator=WindowsDocumentCoordinator(bridge=bridge),
                             session_id=session_id, launcher=launcher)
    except Exception as exc:
        if launcher is not None:
            try:
                launcher.close()
            except Exception:
                pass
        raise SessionStartupError("the Excel Action Session could not be constructed") from exc
