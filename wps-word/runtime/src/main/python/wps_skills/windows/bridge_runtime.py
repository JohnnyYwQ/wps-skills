"""Lazy construction of the one Task-owned Windows WPS bridge."""

import os
from pathlib import Path
import threading

from wps_skills.core import timing

from wps_skills.windows.powershell_bridge import (
    JsonLineBridgeTransport,
    PowerShellBridge,
)
from wps_skills.windows.owned_process import ProcessOwnershipUnavailable
from wps_skills.windows.bridge_types import BackendActionFailure


def windows_powershell_executable():
    windows = Path(os.environ.get("WINDIR", r"C:\Windows"))
    # WPS exposes an out-of-process COM server. Prefer the native host: on
    # 64-bit Windows it avoids the large WOW64 PowerShell startup penalty.
    executable = (
        windows
        / "System32"
        / "WindowsPowerShell"
        / "v1.0"
        / "powershell.exe"
    )
    if executable.is_file():
        return str(executable)
    raise FileNotFoundError("Windows PowerShell could not be located")


class LazyWindowsBridge:
    """Start exactly one owned bridge on the first document acquisition step."""

    def __init__(
        self,
        *,
        launcher,
        script_path,
        executable_factory=windows_powershell_executable,
        transport_factory=JsonLineBridgeTransport,
        bridge_factory=PowerShellBridge,
        debug_close_created_document=False,
    ):
        if not isinstance(debug_close_created_document, bool):
            raise TypeError("debug document cleanup flag must be Boolean")
        self._launcher = launcher
        self._script_path = Path(script_path).resolve()
        self._executable_factory = executable_factory
        self._transport_factory = transport_factory
        self._bridge_factory = bridge_factory
        self._debug_close_created_document = debug_close_created_document
        self._bridge = None
        self._lock = threading.Lock()
        self._closed = False

    def _get_bridge(self):
        with self._lock:
            if self._closed:
                raise BackendActionFailure(
                    outcome="unknown",
                    code="RESPONSE_LOST",
                    message="The WPS bridge is closed",
                    binding_disposition="unprovable",
                )
            if self._bridge is None:
                try:
                    command = [
                        self._executable_factory(),
                        "-NoLogo",
                        "-NoProfile",
                        "-WindowStyle",
                        "Hidden",
                        "-NonInteractive",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-File",
                        str(self._script_path),
                    ]
                    if self._debug_close_created_document:
                        command.append("-DebugCloseCreatedDocument")
                    with timing.span("bridge.process_start"):
                        process = self._launcher.start_bridge(command)
                    self._bridge = self._bridge_factory(
                        transport=self._transport_factory(process=process)
                    )
                except (OSError, ProcessOwnershipUnavailable) as exc:
                    raise BackendActionFailure(
                        outcome="failed",
                        code="DOCUMENT_BINDING_UNAVAILABLE",
                        message="The owned WPS bridge could not be started",
                        binding_disposition="unchanged",
                    ) from exc
            return self._bridge

    def execute(self, operation, arguments, context):
        return self._get_bridge().execute(operation, arguments, context)

    def close(self):
        with self._lock:
            self._closed = True
            bridge = self._bridge
        if bridge is None:
            return True
        return bridge.close()
