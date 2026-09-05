"""Assemble one Windows Word Action Session with application-owned resources."""

from pathlib import Path

from wps_skills.host.session_host import SessionStartupError


BRIDGE_SCRIPT = (
    Path(__file__).resolve().parents[4]
    / "resources" / "wps_skills" / "word" / "windows" / "word_bridge.ps1"
)


def build_session(
    *,
    session_id,
    debug_close_created_document=False,
):
    from wps_skills.core.action_session import ActionSession
    from wps_skills.windows.powershell_bridge import JsonLineBridgeTransport
    from wps_skills.windows.document_coordinator import WindowsDocumentCoordinator
    from wps_skills.windows.owned_process import WindowsOwnedProcessLauncher
    from wps_skills.word.windows.backend import WindowsWordBackend
    from wps_skills.windows.bridge_runtime import LazyWindowsBridge
    from wps_skills.word.adapter import WordAdapter
    from wps_skills.word.contracts import WORD_PRODUCTION_CONTRACT_SET
    from wps_skills.word.handlers import WORD_HANDLERS

    launcher = None
    try:
        launcher = WindowsOwnedProcessLauncher()
        bridge = LazyWindowsBridge(
            launcher=launcher,
            script_path=BRIDGE_SCRIPT,
            transport_factory=JsonLineBridgeTransport,
            debug_close_created_document=debug_close_created_document,
        )
        backend = WindowsWordBackend(bridge=bridge)
        adapter = WordAdapter(
            backend=backend,
            handlers=WORD_HANDLERS,
            contracts=WORD_PRODUCTION_CONTRACT_SET,
        )
        return ActionSession(
            application="word",
            contracts=WORD_PRODUCTION_CONTRACT_SET,
            adapter=adapter,
            coordinator=WindowsDocumentCoordinator(bridge=bridge),
            session_id=session_id,
            launcher=launcher,
        )
    except Exception as exc:
        if launcher is not None:
            try:
                launcher.close()
            except Exception:
                pass
        raise SessionStartupError(
            "the production Word Action Session could not be constructed"
        ) from exc
