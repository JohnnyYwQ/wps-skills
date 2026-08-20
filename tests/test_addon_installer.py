"""Black-box tests for WPS Add-in installation and readiness."""

import json
import os
import socket
import subprocess
import sys
import tempfile
import threading
import time
import unittest
import xml.etree.ElementTree as ElementTree
from pathlib import Path
from urllib.error import URLError
from urllib.request import Request, urlopen


REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
RUNNER = REPOSITORY_ROOT / "skills" / "wps-office" / "scripts" / "wps.py"
POLL_URL = "http://127.0.0.1:58891/poll"
ACK_URL = "http://127.0.0.1:58891/ack"
RESULT_URL = "http://127.0.0.1:58891/result"
REQUEST_ID_HEADER = "X-WPS-Request-ID"


def run_runner(
    operation,
    profile,
    platform_name="linux",
    architecture="x86_64",
    options=None,
    wps_running=None,
    environment_overrides=None,
):
    environment = os.environ.copy()
    environment.update(
        {
            "HOME": str(profile),
            "XDG_CONFIG_HOME": str(profile / ".config"),
            "XDG_DATA_HOME": str(profile / ".local" / "share"),
            "APPDATA": str(profile / "AppData" / "Roaming"),
            "WPS_SKILL_TEST_PLATFORM": platform_name,
            "WPS_SKILL_TEST_ARCHITECTURE": architecture,
        }
    )
    if wps_running is not None:
        environment["WPS_SKILL_TEST_WPS_RUNNING"] = (
            "1" if wps_running else "0"
        )
    if environment_overrides is not None:
        environment.update(environment_overrides)
    runner_argv = [sys.executable, str(RUNNER), operation]
    if options is not None:
        runner_argv.append(json.dumps(options))
    return subprocess.run(
        runner_argv,
        cwd=REPOSITORY_ROOT,
        env=environment,
        text=True,
        capture_output=True,
        timeout=5,
        check=False,
    )


class FakeReadinessAddin:
    """Poll the public readiness protocol as an installed WPS Add-in."""

    def __init__(self, auth_token, install_digest, result=None):
        self.auth_token = auth_token
        self.install_digest = install_digest
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
        headers = {
            "Authorization": "Bearer {0}".format(self.auth_token),
            "Content-Type": "application/json",
        }
        if method == "POST" and self.action_request is not None:
            headers[REQUEST_ID_HEADER] = self.action_request["requestId"]
        return Request(
            url,
            data=data,
            headers=headers,
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
                with urlopen(
                    self._request(ACK_URL, b"", "POST"), timeout=1
                ) as response:
                    json.load(response)
                addin_result = self.result
                if addin_result is None:
                    addin_result = {
                            "success": True,
                            "message": "pong",
                            "timestamp": 1723852800000,
                            "installDigest": self.install_digest,
                    }
                result = json.dumps(
                    {
                        "requestId": self.action_request["requestId"],
                        "result": addin_result,
                    }
                ).encode("utf-8")
                with urlopen(
                    self._request(RESULT_URL, data=result, method="POST"), timeout=1
                ) as response:
                    json.load(response)
                return
            if not self.stopped.is_set():
                raise AssertionError("Readiness check did not publish ping")
        except BaseException as error:
            self.error = error
        finally:
            self.finished.set()


class AddinInstallerBlackBoxTests(unittest.TestCase):
    def test_fresh_linux_profile_gets_a_user_level_addin(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)

            completed = run_runner("install", profile)

            addons = profile / ".local/share/Kingsoft/wps/jsaddons"
            addon = addons / "wps-office-skill_"
            self.assertEqual(completed.returncode, 0, completed.stderr)
            self.assertEqual(
                json.loads(completed.stdout),
                {
                    "ok": True,
                    "operation": "install",
                    "data": {
                        "platform": "linux",
                        "architecture": "x86_64",
                        "status": "installed",
                        "restart_required": True,
                        "addon_path": str(addon),
                    },
                },
            )
            self.assertEqual(completed.stderr, "")
            self.assertTrue((addon / "main.js").is_file())
            self.assertTrue((addons / "publish.xml").is_file())

    def test_windows_arm64_uses_the_roaming_user_profile(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)

            completed = run_runner(
                "install", profile, platform_name="windows", architecture="ARM64"
            )

            addon = (
                profile
                / "AppData/Roaming/kingsoft/wps/jsaddons/wps-office-skill_"
            )
            result = json.loads(completed.stdout)
            self.assertEqual(completed.returncode, 0, completed.stderr)
            self.assertEqual(result["data"]["platform"], "windows")
            self.assertEqual(result["data"]["architecture"], "arm64")
            self.assertEqual(result["data"]["addon_path"], str(addon))
            self.assertTrue((addon / "main.js").is_file())

    def test_macos_arm64_uses_the_wps_container_profile(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            commands = profile / "commands"
            commands.mkdir()
            xattr_log = profile / "xattr.log"
            xattr = commands / "xattr"
            xattr.write_text(
                "#!/bin/sh\nprintf '%s\\n' \"$@\" > \"$WPS_SKILL_XATTR_LOG\"\n",
                encoding="utf-8",
            )
            xattr.chmod(0o755)

            completed = run_runner(
                "install",
                profile,
                platform_name="darwin",
                architecture="arm64",
                environment_overrides={
                    "PATH": str(commands) + os.pathsep + os.environ["PATH"],
                    "WPS_SKILL_XATTR_LOG": str(xattr_log),
                },
            )

            addon = (
                profile
                / "Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps"
                / "jsaddons/wps-office-skill_"
            )
            result = json.loads(completed.stdout)
            self.assertEqual(completed.returncode, 0, completed.stderr)
            self.assertEqual(result["data"]["platform"], "macos")
            self.assertEqual(result["data"]["architecture"], "arm64")
            self.assertEqual(result["data"]["addon_path"], str(addon))
            self.assertTrue((addon / "main.js").is_file())
            self.assertTrue((addon.parent / "publish.xml").is_file())
            self.assertEqual(
                xattr_log.read_text(encoding="utf-8").splitlines(),
                ["-dr", "com.apple.quarantine", str(addon)],
            )

    def test_macos_install_fails_when_quarantine_cannot_be_cleared(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            commands = profile / "commands"
            commands.mkdir()
            xattr = commands / "xattr"
            xattr.write_text("#!/bin/sh\nexit 23\n", encoding="utf-8")
            xattr.chmod(0o755)

            completed = run_runner(
                "install",
                profile,
                platform_name="darwin",
                architecture="arm64",
                environment_overrides={
                    "PATH": str(commands) + os.pathsep + os.environ["PATH"],
                },
            )

            self.assertEqual(completed.returncode, 1)
            result = json.loads(completed.stdout)
            self.assertEqual(
                result["error"]["code"], "MACOS_QUARANTINE_CLEAR_FAILED"
            )

    def test_macos_install_accepts_an_already_clear_quarantine_attribute(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            commands = profile / "commands"
            commands.mkdir()
            xattr = commands / "xattr"
            xattr.write_text(
                "#!/bin/sh\necho 'No such xattr' >&2\nexit 1\n",
                encoding="utf-8",
            )
            xattr.chmod(0o755)

            completed = run_runner(
                "install",
                profile,
                platform_name="darwin",
                architecture="arm64",
                environment_overrides={
                    "PATH": str(commands) + os.pathsep + os.environ["PATH"],
                },
            )

            self.assertEqual(completed.returncode, 0, completed.stderr)
            self.assertEqual(
                json.loads(completed.stdout)["data"]["status"], "installed"
            )

    def test_installer_recognizes_linux_arm64_and_windows_x86_64(self):
        cases = (
            ("linux", "aarch64", "arm64"),
            ("windows", "AMD64", "x86_64"),
        )
        for platform_name, detected_architecture, expected_architecture in cases:
            with self.subTest(
                platform=platform_name, architecture=detected_architecture
            ):
                with tempfile.TemporaryDirectory() as temporary_directory:
                    profile = Path(temporary_directory)

                    completed = run_runner(
                        "install",
                        profile,
                        platform_name=platform_name,
                        architecture=detected_architecture,
                    )

                    self.assertEqual(completed.returncode, 0, completed.stderr)
                    result = json.loads(completed.stdout)
                    self.assertEqual(result["data"]["platform"], platform_name)
                    self.assertEqual(
                        result["data"]["architecture"], expected_architecture
                    )

    def test_linux_uses_an_existing_compatible_wps_profile(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            compatible = profile / ".kingsoft/wps/jsaddons"
            compatible.mkdir(parents=True)

            completed = run_runner("install", profile)

            self.assertEqual(completed.returncode, 0, completed.stderr)
            result = json.loads(completed.stdout)
            self.assertEqual(
                result["data"]["addon_path"],
                str(compatible / "wps-office-skill_"),
            )
            self.assertFalse((profile / ".local/share/Kingsoft").exists())

    def test_repeated_install_is_idempotent_and_keeps_a_protected_credential(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)

            first = run_runner("install", profile)
            addon = profile / ".local/share/Kingsoft/wps/jsaddons/wps-office-skill_"
            config = profile / ".config/wps-office-skill/config.json"
            registry = addon.parent / "publish.xml"
            first_mtimes = {
                "addon": (addon / "main.js").stat().st_mtime_ns,
                "config": config.stat().st_mtime_ns,
                "registry": registry.stat().st_mtime_ns,
            }

            second = run_runner("install", profile)

            self.assertEqual(first.returncode, 0, first.stderr)
            self.assertEqual(second.returncode, 0, second.stderr)
            self.assertEqual(json.loads(second.stdout)["data"]["status"], "current")
            self.assertFalse(
                json.loads(second.stdout)["data"]["restart_required"]
            )
            self.assertEqual(
                first_mtimes,
                {
                    "addon": (addon / "main.js").stat().st_mtime_ns,
                    "config": config.stat().st_mtime_ns,
                    "registry": registry.stat().st_mtime_ns,
                },
            )
            credential = json.loads(config.read_text(encoding="utf-8"))
            self.assertEqual(credential["version"], 1)
            self.assertGreaterEqual(len(credential["auth_token"]), 32)
            runtime_config = (addon / "wps-skill-config.js").read_text(
                encoding="utf-8"
            )
            self.assertIn(credential["auth_token"], runtime_config)
            if os.name != "nt":
                self.assertEqual(config.stat().st_mode & 0o777, 0o600)

    def test_install_preserves_an_existing_addin_registration(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            addons = profile / ".local/share/Kingsoft/wps/jsaddons"
            addons.mkdir(parents=True)
            registry = addons / "publish.xml"
            original = (
                '<?xml version="1.0" encoding="UTF-8"?>\n'
                "<jsplugins>\n"
                '  <jsplugin name="calendar" type="wps" '
                'url="calendar_/" enable="enable_dev"/>\n'
                "</jsplugins>\n"
            )
            registry.write_text(original, encoding="utf-8")

            completed = run_runner("install", profile)

            self.assertEqual(completed.returncode, 0, completed.stderr)
            registrations = {
                entry.get("name"): entry.attrib
                for entry in ElementTree.parse(registry).getroot()
            }
            self.assertEqual(
                registrations["calendar"],
                {
                    "name": "calendar",
                    "type": "wps",
                    "url": "calendar_/",
                    "enable": "enable_dev",
                },
            )
            self.assertEqual(
                registrations["wps-office-skill"],
                {
                    "name": "wps-office-skill",
                    "type": "wps,et,wpp",
                    "url": "wps-office-skill_/",
                    "enable": "enable_dev",
                },
            )
            self.assertEqual(
                (addons / "publish.xml.bak").read_text(encoding="utf-8"),
                original,
            )

    def test_upgrade_atomically_replaces_an_older_addin(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            first = run_runner("install", profile)
            addon = profile / ".local/share/Kingsoft/wps/jsaddons/wps-office-skill_"
            metadata = addon / ".wps-skill-install.json"
            metadata.write_text(
                '{"version":1,"source_digest":"old-release"}\n',
                encoding="utf-8",
            )
            (addon / "obsolete.js").write_text("stale", encoding="utf-8")

            upgraded = run_runner("install", profile)

            self.assertEqual(first.returncode, 0, first.stderr)
            self.assertEqual(upgraded.returncode, 0, upgraded.stderr)
            self.assertEqual(json.loads(upgraded.stdout)["data"]["status"], "updated")
            self.assertTrue(
                json.loads(upgraded.stdout)["data"]["restart_required"]
            )
            self.assertFalse((addon / "obsolete.js").exists())
            self.assertNotEqual(
                json.loads(metadata.read_text(encoding="utf-8"))["source_digest"],
                "old-release",
            )
            self.assertEqual(
                list(addon.parent.glob(".wps-office-skill-stage-*")), []
            )
            self.assertEqual(
                list(addon.parent.glob(".wps-office-skill-backup-*")), []
            )

    def test_install_repairs_a_damaged_addin_even_when_metadata_is_unchanged(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            first = run_runner("install", profile)
            addon = profile / ".local/share/Kingsoft/wps/jsaddons/wps-office-skill_"
            main_script = addon / "main.js"
            main_script.write_text("damaged", encoding="utf-8")

            repaired = run_runner("install", profile)

            self.assertEqual(first.returncode, 0, first.stderr)
            self.assertEqual(repaired.returncode, 0, repaired.stderr)
            self.assertEqual(json.loads(repaired.stdout)["data"]["status"], "updated")
            self.assertNotEqual(main_script.read_text(encoding="utf-8"), "damaged")

    def test_malformed_registry_stops_without_replacing_user_data(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            addons = profile / ".local/share/Kingsoft/wps/jsaddons"
            addons.mkdir(parents=True)
            registry = addons / "publish.xml"
            malformed = "<jsplugins><jsplugin"
            registry.write_text(malformed, encoding="utf-8")

            completed = run_runner("install", profile)

            self.assertEqual(completed.returncode, 1)
            result = json.loads(completed.stdout)
            self.assertEqual(result["operation"], "install")
            self.assertEqual(result["error"]["code"], "INVALID_ADDIN_REGISTRY")
            self.assertFalse(result["error"]["retryable"])
            self.assertEqual(completed.stderr, "")
            self.assertEqual(registry.read_text(encoding="utf-8"), malformed)
            self.assertFalse((addons / "wps-office-skill_").exists())
            self.assertFalse((addons / "publish.xml.bak").exists())

    def test_corrupted_auth_config_is_not_silently_replaced(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            config = profile / ".config/wps-office-skill/config.json"
            config.parent.mkdir(parents=True)
            config.write_text("{broken", encoding="utf-8")

            completed = run_runner("install", profile)

            self.assertEqual(completed.returncode, 1)
            result = json.loads(completed.stdout)
            self.assertEqual(result["operation"], "install")
            self.assertEqual(result["error"]["code"], "INVALID_AUTH_CONFIG")
            self.assertEqual(config.read_text(encoding="utf-8"), "{broken")
            self.assertFalse(
                (
                    profile
                    / ".local/share/Kingsoft/wps/jsaddons/wps-office-skill_"
                ).exists()
            )

    def test_unsupported_architecture_does_not_install_anything(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)

            completed = run_runner("install", profile, architecture="riscv64")

            self.assertEqual(completed.returncode, 1)
            result = json.loads(completed.stdout)
            self.assertEqual(result["error"]["code"], "UNSUPPORTED_ARCHITECTURE")
            self.assertFalse((profile / ".local").exists())

    def test_first_readiness_check_installs_and_requests_a_wps_restart(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)

            completed = run_runner("check", profile)

            self.assertEqual(completed.returncode, 0, completed.stderr)
            self.assertEqual(
                json.loads(completed.stdout),
                {
                    "ok": True,
                    "operation": "check",
                    "data": {
                        "status": "restart_required",
                        "ready": False,
                        "restart_required": True,
                        "platform": "linux",
                        "architecture": "x86_64",
                    },
                },
            )
            self.assertEqual(completed.stderr, "")

    def test_readiness_distinguishes_wps_not_running_after_install(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            installed = run_runner("install", profile)

            completed = run_runner("check", profile, wps_running=False)

            self.assertEqual(installed.returncode, 0, installed.stderr)
            self.assertEqual(completed.returncode, 0, completed.stderr)
            data = json.loads(completed.stdout)["data"]
            self.assertEqual(data["status"], "wps_not_running")
            self.assertFalse(data["ready"])
            self.assertFalse(data["restart_required"])

    def test_restart_required_persists_until_the_loaded_addin_matches(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            installed = run_runner("install", profile)

            completed = run_runner(
                "check",
                profile,
                options={"timeout_ms": 50},
                wps_running=True,
            )

            self.assertEqual(installed.returncode, 0, installed.stderr)
            self.assertEqual(completed.returncode, 0, completed.stderr)
            data = json.loads(completed.stdout)["data"]
            self.assertEqual(data["status"], "restart_required")
            self.assertFalse(data["ready"])
            self.assertTrue(data["restart_required"])

    def test_readiness_distinguishes_a_running_wps_without_the_addin(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            installed = run_runner("install", profile)

            config = json.loads(
                (profile / ".config/wps-office-skill/config.json").read_text(
                    encoding="utf-8"
                )
            )
            digest = json.loads(
                (
                    profile
                    / ".local/share/Kingsoft/wps/jsaddons/wps-office-skill_/.wps-skill-install.json"
                ).read_text(encoding="utf-8")
            )["source_digest"]
            addin = FakeReadinessAddin(config["auth_token"], digest)
            addin.start()
            acknowledged = run_runner(
                "check", profile, options={"timeout_ms": 2000}, wps_running=True
            )

            completed = run_runner(
                "check",
                profile,
                options={"timeout_ms": 50},
                wps_running=True,
            )

            self.assertEqual(installed.returncode, 0, installed.stderr)
            self.assertTrue(
                addin.finished.wait(1),
                "Fake Add-in did not finish: " + completed.stdout + completed.stderr,
            )
            if addin.error:
                raise addin.error
            self.assertEqual(
                json.loads(acknowledged.stdout)["data"]["status"], "ready"
            )
            self.assertEqual(completed.returncode, 0, completed.stderr)
            data = json.loads(completed.stdout)["data"]
            self.assertEqual(data["status"], "addin_unavailable")
            self.assertFalse(data["ready"])
            self.assertFalse(data["restart_required"])

    def test_authenticated_ping_reports_the_addin_ready(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            installed = run_runner("install", profile)
            config = json.loads(
                (profile / ".config/wps-office-skill/config.json").read_text(
                    encoding="utf-8"
                )
            )
            digest = json.loads(
                (
                    profile
                    / ".local/share/Kingsoft/wps/jsaddons/wps-office-skill_/.wps-skill-install.json"
                ).read_text(encoding="utf-8")
            )["source_digest"]
            addin = FakeReadinessAddin(config["auth_token"], digest)
            addin.start()

            completed = run_runner(
                "check",
                profile,
                options={"timeout_ms": 2000},
                wps_running=True,
            )

            self.assertEqual(installed.returncode, 0, installed.stderr)
            self.assertTrue(addin.finished.wait(1), "Fake Add-in did not finish")
            if addin.error:
                raise addin.error
            self.assertEqual(completed.returncode, 0, completed.stderr)
            data = json.loads(completed.stdout)["data"]
            self.assertEqual(data["status"], "ready")
            self.assertTrue(data["ready"])
            self.assertFalse(data["restart_required"])
            self.assertEqual(addin.action_request["action"], "ping")
            self.assertEqual(addin.action_request["params"], {})

    def test_ping_with_a_different_credential_cannot_report_ready(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            installed = run_runner("install", profile)
            config = json.loads(
                (profile / ".config/wps-office-skill/config.json").read_text(
                    encoding="utf-8"
                )
            )
            digest = json.loads(
                (
                    profile
                    / ".local/share/Kingsoft/wps/jsaddons/wps-office-skill_/.wps-skill-install.json"
                ).read_text(encoding="utf-8")
            )["source_digest"]
            trusted_addin = FakeReadinessAddin(config["auth_token"], digest)
            trusted_addin.start()
            acknowledged = run_runner(
                "check", profile, options={"timeout_ms": 2000}, wps_running=True
            )
            self.assertTrue(
                trusted_addin.finished.wait(1), "Trusted Fake Add-in did not finish"
            )
            if trusted_addin.error:
                raise trusted_addin.error
            self.assertEqual(
                json.loads(acknowledged.stdout)["data"]["status"], "ready"
            )

            addin = FakeReadinessAddin(
                "wrong-credential-that-is-long-enough", digest
            )
            addin.start()

            completed = run_runner(
                "check",
                profile,
                options={"timeout_ms": 100},
                wps_running=True,
            )
            addin.stop()

            self.assertEqual(installed.returncode, 0, installed.stderr)
            self.assertTrue(addin.finished.wait(1), "Fake Add-in did not stop")
            if addin.error:
                raise addin.error
            self.assertEqual(completed.returncode, 0, completed.stderr)
            data = json.loads(completed.stdout)["data"]
            self.assertEqual(data["status"], "addin_unavailable")
            self.assertFalse(data["ready"])
            self.assertIsNone(addin.action_request)

    def test_invoke_performs_first_use_install_before_delivering_an_action(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)

            completed = run_runner(
                "invoke",
                profile,
                options={"action": "ping", "params": {}, "timeout_ms": 50},
                wps_running=True,
            )

            self.assertEqual(completed.returncode, 1)
            self.assertEqual(
                json.loads(completed.stdout),
                {
                    "ok": False,
                    "action": "ping",
                    "error": {
                        "code": "WPS_RESTART_REQUIRED",
                        "message": (
                            "WPS Add-in was installed or updated; restart WPS "
                            "Office before retrying"
                        ),
                        "retryable": True,
                    },
                },
            )

    def test_invoke_reports_when_wps_is_not_running(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            installed = run_runner("install", profile)

            completed = run_runner(
                "invoke",
                profile,
                options={"action": "ping", "params": {}, "timeout_ms": 50},
                wps_running=False,
            )

            self.assertEqual(installed.returncode, 0, installed.stderr)
            self.assertEqual(completed.returncode, 1)
            result = json.loads(completed.stdout)
            self.assertEqual(result["error"]["code"], "WPS_NOT_RUNNING")
            self.assertTrue(result["error"]["retryable"])

    def test_occupied_readiness_port_reports_addin_unavailable(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            installed = run_runner("install", profile)
            config = json.loads(
                (profile / ".config/wps-office-skill/config.json").read_text(
                    encoding="utf-8"
                )
            )
            digest = json.loads(
                (
                    profile
                    / ".local/share/Kingsoft/wps/jsaddons/wps-office-skill_/.wps-skill-install.json"
                ).read_text(encoding="utf-8")
            )["source_digest"]
            addin = FakeReadinessAddin(config["auth_token"], digest)
            addin.start()
            acknowledged = run_runner(
                "check", profile, options={"timeout_ms": 2000}, wps_running=True
            )
            self.assertTrue(addin.finished.wait(1), "Fake Add-in did not finish")
            if addin.error:
                raise addin.error

            with socket.socket() as occupied_port:
                occupied_port.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                occupied_port.bind(("127.0.0.1", 58891))
                occupied_port.listen(1)
                completed = run_runner(
                    "check",
                    profile,
                    options={"timeout_ms": 50},
                    wps_running=True,
                )

            self.assertEqual(installed.returncode, 0, installed.stderr)
            self.assertEqual(
                json.loads(acknowledged.stdout)["data"]["status"], "ready"
            )
            self.assertEqual(
                json.loads(completed.stdout)["data"]["status"],
                "addin_unavailable",
            )
            self.assertEqual(
                json.loads(completed.stdout)["data"]["error"],
                {
                    "code": "PORT_IN_USE",
                    "message": "Loopback port 58891 is already in use",
                    "retryable": True,
                },
            )

    def test_invalid_ping_is_not_misclassified_as_restart_required(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            profile = Path(temporary_directory)
            run_runner("install", profile)
            config = json.loads(
                (profile / ".config/wps-office-skill/config.json").read_text(
                    encoding="utf-8"
                )
            )
            digest = json.loads(
                (
                    profile
                    / ".local/share/Kingsoft/wps/jsaddons/wps-office-skill_/.wps-skill-install.json"
                ).read_text(encoding="utf-8")
            )["source_digest"]
            trusted_addin = FakeReadinessAddin(config["auth_token"], digest)
            trusted_addin.start()
            run_runner(
                "check", profile, options={"timeout_ms": 2000}, wps_running=True
            )
            self.assertTrue(trusted_addin.finished.wait(1))
            if trusted_addin.error:
                raise trusted_addin.error

            invalid_addin = FakeReadinessAddin(
                config["auth_token"],
                digest,
                result={"success": False, "error": "invalid ping"},
            )
            invalid_addin.start()
            completed = run_runner(
                "check", profile, options={"timeout_ms": 2000}, wps_running=True
            )

            self.assertTrue(invalid_addin.finished.wait(1))
            if invalid_addin.error:
                raise invalid_addin.error
            self.assertEqual(
                json.loads(completed.stdout)["data"]["status"],
                "addin_unavailable",
            )
            self.assertEqual(
                json.loads(completed.stdout)["data"]["error"]["code"],
                "INVALID_RESULT",
            )


if __name__ == "__main__":
    unittest.main()
