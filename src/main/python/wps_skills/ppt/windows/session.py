"""Assemble one Windows Ppt Action Session with application-owned resources."""

from pathlib import Path

from wps_skills.host.session_host import SessionStartupError


BRIDGE_SCRIPT = (
    Path(__file__).resolve().parents[4]
    / "resources" / "wps_skills" / "ppt" / "windows" / "ppt_bridge.ps1"
)


def build_session(*, session_id, contracts=None):
    from wps_skills.core.action_session import ActionSession
    from wps_skills.ppt.adapter import PptAdapter
    from wps_skills.ppt.contracts import PPT_PRODUCTION_CONTRACT_SET
    from wps_skills.windows.bridge_runtime import LazyWindowsBridge
    from wps_skills.windows.document_coordinator import WindowsDocumentCoordinator
    from wps_skills.ppt.windows.backend import WindowsPptBackend
    from wps_skills.windows.owned_process import WindowsOwnedProcessLauncher

    contracts = PPT_PRODUCTION_CONTRACT_SET if contracts is None else contracts
    if contracts is None:
        raise SessionStartupError("Ppt implementation awaits live-WPS admission; no production Ppt Application Contract Set is installed")
    launcher = None
    try:
        launcher = WindowsOwnedProcessLauncher()
        bridge = LazyWindowsBridge(launcher=launcher, script_path=BRIDGE_SCRIPT)
        return ActionSession(application="ppt", contracts=contracts,
                             adapter=PptAdapter(backend=WindowsPptBackend(bridge=bridge), contracts=contracts),
                             coordinator=WindowsDocumentCoordinator(bridge=bridge),
                             session_id=session_id, launcher=launcher)
    except Exception as exc:
        if launcher is not None:
            try:
                launcher.close()
            except Exception:
                pass
        raise SessionStartupError("the Ppt Action Session could not be constructed") from exc
