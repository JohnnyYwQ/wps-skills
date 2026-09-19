"""Windows regression: collecting evidence must not block concurrent log appends.

No WPS, network or user documents. Temporary files only; finishes in seconds.
"""
import argparse
import base64
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time


def literal(value):
    return "'" + str(value).replace("'", "''") + "'"


def main(copier):
    if os.name != "nt":
        raise RuntimeError("Windows file sharing semantics required")
    with tempfile.TemporaryDirectory(prefix="wps-snapshot-sharing-") as temporary:
        root = Path(temporary)
        source = root / "events.jsonl"
        source.write_bytes(b"x" * (16 * 1024 * 1024))
        script = "$ErrorActionPreference='Stop'; . " + literal(copier.resolve()) + "\n"
        script += "[IO.File]::WriteAllText(" + literal(root / "ready") + ", 'ready')\n"
        script += "for ($i=0; $i -lt 20; $i++) { Copy-DiagnosticEvidenceFile -Source " + literal(source) + " -Destination " + literal(root / "snapshot") + " }"
        encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        executable = str(Path(os.environ["WINDIR"]) / "System32/WindowsPowerShell/v1.0/powershell.exe")
        with (root / "stderr.txt").open("w") as stderr:
            process = subprocess.Popen([executable, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
                                       stdout=subprocess.DEVNULL, stderr=stderr)
            deadline = time.monotonic() + 15
            writes = failures = 0
            try:
                while not (root / "ready").exists() and process.poll() is None and time.monotonic() < deadline:
                    time.sleep(0.001)
                while process.poll() is None and time.monotonic() < deadline:
                    try:
                        with source.open("ab") as stream:
                            stream.write(b"\n")
                        writes += 1
                    except PermissionError:
                        failures += 1
                    time.sleep(0.001)
                code = process.wait(timeout=2)
            finally:
                if process.poll() is None:
                    process.kill()
                    process.wait()
        result = {"appendSuccesses": writes, "appendSharingViolations": failures, "copyExitCode": code}
        print(json.dumps(result))
        if code:
            print((root / "stderr.txt").read_text(errors="replace")[:3000])
        if code or failures or not writes:
            print("FAIL: evidence collection interfered with the live log writer")
            return 1
        print("PASS: concurrent evidence copy did not block log appends")
        return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--copier", type=Path, required=True)
    raise SystemExit(main(parser.parse_args().copier))
