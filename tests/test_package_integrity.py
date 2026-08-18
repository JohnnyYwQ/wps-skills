"""Black-box isolation tests for the distributable WPS Skill Package."""

import ast
import importlib.machinery
import json
import os
import shutil
import socket
import subprocess
import sys
import sysconfig
import tempfile
import threading
import time
import unittest
from pathlib import Path
from urllib.error import URLError
from urllib.request import Request, urlopen


REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
PACKAGE_SOURCE = REPOSITORY_ROOT / "skills" / "wps-office"
POLL_URL = "http://127.0.0.1:58891/poll"
RESULT_URL = "http://127.0.0.1:58891/result"


def isolated_environment(profile, audit_directory):
    """Build an environment with no repository Python import path."""
    environment = {
        key: value
        for key, value in os.environ.items()
        if key
        not in {
            "CONDA_PREFIX",
            "PIP_CONFIG_FILE",
            "PYTHONHOME",
            "PYTHONPATH",
            "PYTHONUSERBASE",
            "VIRTUAL_ENV",
        }
    }
    environment.update(
        {
            "HOME": str(profile),
            "XDG_CONFIG_HOME": str(profile / ".config"),
            "XDG_DATA_HOME": str(profile / ".local" / "share"),
            "APPDATA": str(profile / "AppData" / "Roaming"),
            "PYTHONPATH": str(audit_directory),
            "PYTHONDONTWRITEBYTECODE": "1",
            "WPS_SKILL_FORBIDDEN_REPOSITORY": str(REPOSITORY_ROOT),
            "WPS_SKILL_TEST_PLATFORM": "linux",
            "WPS_SKILL_TEST_ARCHITECTURE": "x86_64",
            "WPS_SKILL_TEST_WPS_RUNNING": "1",
        }
    )
    return environment


def run_runner(package, operation, environment, options=None):
    command = [sys.executable, str(package / "scripts" / "wps.py"), operation]
    if options is not None:
        command.append(json.dumps(options))
    return subprocess.run(
        command,
        cwd=package,
        env=environment,
        text=True,
        capture_output=True,
        timeout=5,
        check=False,
    )


def loopback_port_is_available():
    with socket.socket() as candidate:
        candidate.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            candidate.bind(("127.0.0.1", 58891))
        except OSError:
            return False
    return True


def retryable_port_failure(completed):
    """Recognize the Runner's documented, retryable loopback-port failures."""
    try:
        result = json.loads(completed.stdout)
    except ValueError:
        return False
    error = result.get("error")
    if error is None:
        error = result.get("data", {}).get("error")
    return isinstance(error, dict) and error.get("code") in {
        "PORT_IN_USE",
        "PORT_UNAVAILABLE",
    }


def standard_library_modules():
    """Return import roots available in Python 3.9 and newer."""
    module_names = getattr(sys, "stdlib_module_names", None)
    if module_names is not None:
        return set(module_names)

    standard_library = set(sys.builtin_module_names)
    standard_library_directory = Path(sysconfig.get_paths()["stdlib"])
    for path in standard_library_directory.iterdir():
        if path.is_dir() and (path / "__init__.py").is_file():
            standard_library.add(path.name)
        elif path.suffix == ".py":
            standard_library.add(path.stem)
    dynamic_library_directory = standard_library_directory / "lib-dynload"
    if dynamic_library_directory.is_dir():
        for path in dynamic_library_directory.iterdir():
            if any(
                path.name.endswith(suffix)
                for suffix in importlib.machinery.EXTENSION_SUFFIXES
            ):
                standard_library.add(path.name.split(".", 1)[0])
    standard_library.update({"fcntl", "msvcrt"})
    return standard_library


def install_repository_access_guard(audit_directory):
    """Reject source-repository reads in the staged Runner process."""
    audit_directory.mkdir()
    (audit_directory / "sitecustomize.py").write_text(
        """import os
import sys


FORBIDDEN_REPOSITORY = os.path.realpath(
    os.environ[\"WPS_SKILL_FORBIDDEN_REPOSITORY\"]
)


def audit(event, arguments):
    if event != \"open\" or not arguments or not isinstance(arguments[0], str):
        return
    path = os.path.realpath(arguments[0])
    if path == FORBIDDEN_REPOSITORY or path.startswith(
        FORBIDDEN_REPOSITORY + os.sep
    ):
        raise PermissionError(\"The staged package read the source repository\")


sys.addaudithook(audit)
""",
        encoding="utf-8",
    )


class FakeAddin:
    """Use the package's authenticated loopback protocol as a WPS Add-in."""

    def __init__(self, auth_token, result):
        self.auth_token = auth_token
        self.result = result
        self.action_request = None
        self.error = None
        self.finished = threading.Event()
        self.stopped = threading.Event()

    def start(self):
        thread = threading.Thread(target=self._run, daemon=True)
        thread.start()
        return thread

    def stop(self):
        self.stopped.set()

    def _request(self, url, data=None, method="GET"):
        return Request(
            url,
            data=data,
            headers={
                "Authorization": "Bearer {0}".format(self.auth_token),
                "Content-Type": "application/json",
            },
            method=method,
        )

    def _run(self):
        try:
            deadline = time.monotonic() + 5
            while time.monotonic() < deadline and not self.stopped.is_set():
                try:
                    with urlopen(self._request(POLL_URL), timeout=0.25) as response:
                        payload = json.load(response)
                except (OSError, URLError):
                    time.sleep(0.02)
                    continue
                if "actionRequest" not in payload:
                    time.sleep(0.02)
                    continue
                self.action_request = payload["actionRequest"]
                body = json.dumps(
                    {
                        "requestId": self.action_request["requestId"],
                        "result": self.result,
                    }
                ).encode("utf-8")
                with urlopen(
                    self._request(RESULT_URL, body, "POST"), timeout=1
                ) as response:
                    json.load(response)
                return
            if not self.stopped.is_set():
                raise AssertionError("Runner did not publish a WPS Action")
        except BaseException as error:  # surfaced in the test process below
            self.error = error
        finally:
            self.finished.set()


class PackageIntegrityTests(unittest.TestCase):
    def _stage_package(self, directory):
        package = Path(directory) / "isolated-distribution" / "wps-office"
        shutil.copytree(PACKAGE_SOURCE, package)
        return package

    def _run_with_addin(
        self, package, environment, operation, options, result, auth_token
    ):
        deadline = time.monotonic() + 2
        while True:
            addin = FakeAddin(auth_token, result)
            addin.start()
            completed = run_runner(package, operation, environment, options)
            if retryable_port_failure(completed):
                addin.stop()
                self.assertTrue(addin.finished.wait(1), "Fake Add-in did not stop")
                if time.monotonic() >= deadline:
                    self.fail(
                        "Runner could not acquire the loopback port: "
                        + completed.stdout
                        + completed.stderr
                    )
                time.sleep(0.05)
                continue
            self.assertTrue(
                addin.finished.wait(2),
                "Fake Add-in did not finish: "
                + completed.stdout
                + completed.stderr,
            )
            if addin.error:
                raise addin.error
            return completed, addin

    def test_isolated_package_has_one_trigger_and_no_retired_runtime_artifacts(self):
        with tempfile.TemporaryDirectory() as directory:
            package = self._stage_package(directory)
            package_files = [path for path in package.rglob("*") if path.is_file()]

            self.assertEqual(
                [path.relative_to(package).as_posix() for path in package.rglob("SKILL.md")],
                ["SKILL.md"],
            )
            self.assertTrue((package / "scripts/wps.py").is_file())
            self.assertTrue((package / "references/action-manifest.json").is_file())
            self.assertTrue((package / "assets/wps-addin/main.js").is_file())
            self.assertFalse(any(path.is_symlink() for path in package.rglob("*")))

            forbidden_suffixes = {".mjs", ".node", ".ps1", ".psm1", ".ts", ".tsx"}
            self.assertFalse(
                [
                    path.relative_to(package).as_posix()
                    for path in package_files
                    if path.suffix.lower() in forbidden_suffixes
                    or path.name == "package.json"
                ]
            )

            for source_path in package_files:
                if source_path.suffix not in {".js", ".json", ".md", ".py", ".xml"}:
                    continue
                source = source_path.read_text(encoding="utf-8")
                for forbidden_runtime in (
                    "MCP",
                    "Node.js",
                    "npm",
                    "TypeScript",
                    "PowerShell",
                ):
                    with self.subTest(source=source_path, runtime=forbidden_runtime):
                        self.assertNotIn(forbidden_runtime, source)

            standard_library = standard_library_modules() | {
                "__future__",
                "wps_skill",
            }
            for source_path in (package / "scripts").rglob("*.py"):
                for node in ast.walk(ast.parse(source_path.read_text(encoding="utf-8"))):
                    if isinstance(node, ast.Import):
                        imports = [alias.name for alias in node.names]
                    elif isinstance(node, ast.ImportFrom) and node.module:
                        imports = [node.module]
                    else:
                        continue
                    for imported_module in imports:
                        with self.subTest(source=source_path, module=imported_module):
                            self.assertIn(
                                imported_module.split(".", 1)[0], standard_library
                            )

    def test_staged_package_runs_documented_operations_without_repository_resources(self):
        with tempfile.TemporaryDirectory() as directory:
            package = self._stage_package(directory)
            profile = Path(directory) / "profile"
            audit_directory = Path(directory) / "audit"
            install_repository_access_guard(audit_directory)
            environment = isolated_environment(profile, audit_directory)

            initial_check = run_runner(package, "check", environment)

            self.assertEqual(initial_check.returncode, 0, initial_check.stderr)
            self.assertEqual(initial_check.stderr, "")
            self.assertEqual(
                json.loads(initial_check.stdout)["data"]["status"],
                "restart_required",
            )

            installed = run_runner(package, "install", environment)
            self.assertEqual(installed.returncode, 0, installed.stderr)
            self.assertEqual(installed.stderr, "")
            self.assertEqual(
                json.loads(installed.stdout)["data"]["status"], "current"
            )

            addon = (
                profile / ".local/share/Kingsoft/wps/jsaddons/wps-office-skill_"
            )
            config_path = profile / ".config/wps-office-skill/config.json"
            config = json.loads(config_path.read_text(encoding="utf-8"))
            install_digest = json.loads(
                (addon / ".wps-skill-install.json").read_text(encoding="utf-8")
            )["source_digest"]
            self.assertEqual(config_path.stat().st_mode & 0o777, 0o600)
            self.assertIn(config["auth_token"], (addon / "wps-skill-config.js").read_text(encoding="utf-8"))
            for output in (initial_check.stdout, initial_check.stderr, installed.stdout, installed.stderr):
                self.assertNotIn(config["auth_token"], output)
            self.assertFalse(
                [
                    path.name
                    for path in addon.parent.iterdir()
                    if path.name.startswith(".wps-office-skill-")
                    or path.name.startswith(".wps-office-skill-stage-")
                ]
            )

            ready_check, readiness_addin = self._run_with_addin(
                package,
                environment,
                "check",
                {"timeout_ms": 2000},
                {
                    "success": True,
                    "message": "pong",
                    "timestamp": 1723852800000,
                    "installDigest": install_digest,
                },
                config["auth_token"],
            )
            self.assertEqual(ready_check.returncode, 0, ready_check.stderr)
            self.assertEqual(ready_check.stderr, "")
            self.assertEqual(json.loads(ready_check.stdout)["data"]["status"], "ready")
            self.assertEqual(readiness_addin.action_request["action"], "ping")
            self.assertEqual(readiness_addin.action_request["params"], {})

            cases = (
                (
                    {"action": "getCellValue", "params": {"sheet": "Sheet1", "row": 1, "col": 1}},
                    {"success": True, "value": 42, "text": "42", "formula": ""},
                    {"value": 42, "text": "42", "formula": ""},
                ),
                (
                    {"action": "insertText", "params": {"text": "Quarterly report", "position": "end"}},
                    {"success": True, "position": "end", "textLength": 16},
                    {"position": "end", "textLength": 16},
                ),
                (
                    {"action": "deleteSlide", "params": {"slideIndex": 2}, "confirmed": True},
                    {"success": True, "deleted": 2},
                    {"deleted": 2},
                ),
            )
            for request, addin_result, expected_data in cases:
                with self.subTest(action=request["action"]):
                    completed, addin = self._run_with_addin(
                        package,
                        environment,
                        "invoke",
                        {**request, "timeout_ms": 2000},
                        addin_result,
                        config["auth_token"],
                    )
                    self.assertEqual(addin.action_request["action"], request["action"])
                    self.assertEqual(addin.action_request["params"], request["params"])
                    self.assertEqual(completed.returncode, 0, completed.stderr)
                    self.assertEqual(completed.stderr, "")
                    self.assertEqual(
                        json.loads(completed.stdout),
                        {
                            "ok": True,
                            "action": request["action"],
                            "data": expected_data,
                        },
                    )

            self.assertTrue(loopback_port_is_available())


if __name__ == "__main__":
    unittest.main()
