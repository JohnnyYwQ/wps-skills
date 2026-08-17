"""Install the bundled WPS Add-in into the current user's profile."""

import hashlib
import json
import os
import platform
import secrets
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ElementTree
from pathlib import Path


ADDIN_NAME = "wps-office-skill"
ADDIN_DIRECTORY_NAME = ADDIN_NAME + "_"
SKILL_ROOT = Path(__file__).resolve().parents[2]
ADDIN_SOURCE = SKILL_ROOT / "assets" / "wps-addin"
INSTALL_METADATA = ".wps-skill-install.json"
RUNTIME_CONFIG = "wps-skill-config.js"


class AddinInstallError(Exception):
    def __init__(self, code, message, retryable=False):
        super().__init__(message)
        self.code = code
        self.message = message
        self.retryable = retryable


def _platform_name():
    detected = os.environ.get("WPS_SKILL_TEST_PLATFORM", sys.platform)
    if detected.startswith("linux"):
        return "linux"
    if detected in ("win32", "windows"):
        return "windows"
    raise AddinInstallError(
        "UNSUPPORTED_PLATFORM", "Unsupported platform: {0}".format(detected)
    )


def _architecture():
    detected = os.environ.get(
        "WPS_SKILL_TEST_ARCHITECTURE", platform.machine()
    ).lower()
    if detected in ("x86_64", "amd64"):
        return "x86_64"
    if detected in ("aarch64", "arm64"):
        return "arm64"
    raise AddinInstallError(
        "UNSUPPORTED_ARCHITECTURE",
        "Unsupported architecture: {0}".format(detected),
    )


def _addons_directory(platform_name):
    if platform_name == "windows":
        return Path(os.environ["APPDATA"]) / "kingsoft" / "wps" / "jsaddons"
    data_home = Path(
        os.environ.get("XDG_DATA_HOME", str(Path.home() / ".local" / "share"))
    )
    preferred = data_home / "Kingsoft" / "wps" / "jsaddons"
    compatible = Path.home() / ".kingsoft" / "wps" / "jsaddons"
    if not preferred.exists() and compatible.exists():
        return compatible
    return preferred


def _config_path(platform_name):
    if platform_name == "windows":
        config_root = Path(os.environ["APPDATA"])
    else:
        config_root = Path(
            os.environ.get("XDG_CONFIG_HOME", str(Path.home() / ".config"))
        )
    return config_root / "wps-office-skill" / "config.json"


def _write_private_text(path, content):
    path.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
    if os.name != "nt":
        path.parent.chmod(0o700)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=".{0}.".format(path.name), dir=str(path.parent)
    )
    temporary_path = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as temporary_file:
            temporary_file.write(content)
            temporary_file.flush()
            os.fsync(temporary_file.fileno())
        if os.name != "nt":
            temporary_path.chmod(0o600)
        os.replace(str(temporary_path), str(path))
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def _credential(platform_name):
    config_path = _config_path(platform_name)
    if config_path.exists():
        try:
            with config_path.open(encoding="utf-8") as config_file:
                config = json.load(config_file)
        except (OSError, TypeError, ValueError) as error:
            raise AddinInstallError(
                "INVALID_AUTH_CONFIG",
                "The WPS Skill authentication config is invalid: {0}".format(error),
            )
        if (
            not isinstance(config, dict)
            or config.get("version") != 1
            or not isinstance(config.get("auth_token"), str)
            or len(config["auth_token"]) < 32
        ):
            raise AddinInstallError(
                "INVALID_AUTH_CONFIG",
                "The WPS Skill authentication config is invalid",
            )
        if os.name != "nt":
            config_path.chmod(0o600)
        return config["auth_token"]

    config = {"version": 1, "auth_token": secrets.token_urlsafe(32)}
    _write_private_text(
        config_path,
        json.dumps(config, ensure_ascii=False, separators=(",", ":")) + "\n",
    )
    return config["auth_token"]


def auth_token(platform_name):
    """Return the installed credential without exposing it in process output."""
    return _credential(platform_name)


def wps_is_running(platform_name):
    test_value = os.environ.get("WPS_SKILL_TEST_WPS_RUNNING")
    if test_value is not None:
        return test_value == "1"
    process_names = {"wps", "wps.exe", "et", "et.exe", "wpp", "wpp.exe", "wpsoffice", "wpsoffice.exe"}
    if platform_name == "linux":
        proc = Path("/proc")
        if not proc.is_dir():
            return False
        for process_directory in proc.iterdir():
            if not process_directory.name.isdigit():
                continue
            try:
                command = (process_directory / "comm").read_text(
                    encoding="utf-8", errors="ignore"
                ).strip().lower()
            except OSError:
                continue
            if command in process_names:
                return True
        return False
    try:
        completed = subprocess.run(
            ["tasklist", "/FO", "CSV", "/NH"],
            text=True,
            capture_output=True,
            timeout=5,
            check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return False
    first_fields = {
        line.split(",", 1)[0].strip().strip('"').lower()
        for line in completed.stdout.splitlines()
    }
    return bool(first_fields & process_names)


def _source_digest():
    digest = hashlib.sha256()
    for source_path in sorted(path for path in ADDIN_SOURCE.rglob("*") if path.is_file()):
        relative_path = source_path.relative_to(ADDIN_SOURCE).as_posix()
        digest.update(relative_path.encode("utf-8"))
        digest.update(b"\0")
        with source_path.open("rb") as source_file:
            for chunk in iter(lambda: source_file.read(65536), b""):
                digest.update(chunk)
    return digest.hexdigest()


def _runtime_config_content(auth_token):
    return "var WPS_SKILL_AUTH_TOKEN = {0};\n".format(json.dumps(auth_token))


def _installed_digest(addon_path):
    try:
        with (addon_path / INSTALL_METADATA).open(encoding="utf-8") as metadata_file:
            metadata = json.load(metadata_file)
        return metadata.get("source_digest")
    except (OSError, TypeError, ValueError):
        return None


def _addon_is_current(addon_path, source_digest, runtime_config):
    if not addon_path.is_dir() or _installed_digest(addon_path) != source_digest:
        return False
    try:
        return (addon_path / RUNTIME_CONFIG).read_text(
            encoding="utf-8"
        ) == runtime_config
    except OSError:
        return False


def _replace_addon(addons_directory, addon_path, source_digest, runtime_config):
    staging_root = Path(
        tempfile.mkdtemp(prefix=".wps-office-skill-stage-", dir=str(addons_directory))
    )
    staged_addon = staging_root / ADDIN_DIRECTORY_NAME
    backup = addons_directory / ".wps-office-skill-backup-{0}".format(
        secrets.token_hex(8)
    )
    moved_existing = False
    try:
        shutil.copytree(str(ADDIN_SOURCE), str(staged_addon))
        _write_private_text(staged_addon / RUNTIME_CONFIG, runtime_config)
        _write_private_text(
            staged_addon / INSTALL_METADATA,
            json.dumps(
                {"version": 1, "source_digest": source_digest},
                separators=(",", ":"),
            )
            + "\n",
        )
        if addon_path.exists():
            os.replace(str(addon_path), str(backup))
            moved_existing = True
        try:
            os.replace(str(staged_addon), str(addon_path))
        except BaseException:
            if moved_existing:
                os.replace(str(backup), str(addon_path))
            raise
        if moved_existing:
            shutil.rmtree(str(backup))
    finally:
        shutil.rmtree(str(staging_root), ignore_errors=True)


def _load_registry(registry_path):
    if not registry_path.exists():
        return ElementTree.ElementTree(ElementTree.Element("jsplugins"))
    try:
        return ElementTree.parse(str(registry_path))
    except (OSError, ElementTree.ParseError) as error:
        raise AddinInstallError(
            "INVALID_ADDIN_REGISTRY",
            "The WPS Add-in registry is invalid: {0}".format(error),
        )


def _merge_registration(registry_path):
    tree = _load_registry(registry_path)
    root = tree.getroot()
    if root.tag.split("}")[-1] != "jsplugins":
        raise AddinInstallError(
            "INVALID_ADDIN_REGISTRY",
            "The WPS Add-in registry root must be jsplugins",
        )
    expected = {
        "name": ADDIN_NAME,
        "type": "wps,et,wpp",
        "url": ADDIN_DIRECTORY_NAME + "/",
        "enable": "enable_dev",
    }
    registrations = [
        child
        for child in root
        if child.tag.split("}")[-1] == "jsplugin"
        and child.get("name") == ADDIN_NAME
    ]
    if len(registrations) == 1 and all(
        registrations[0].get(key) == value for key, value in expected.items()
    ):
        return False

    if registrations:
        registration = registrations[0]
        for duplicate in registrations[1:]:
            root.remove(duplicate)
        registration.attrib.update(expected)
    else:
        ElementTree.SubElement(root, "jsplugin", expected)

    if registry_path.exists():
        shutil.copy2(str(registry_path), str(registry_path) + ".bak")
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=".publish.", suffix=".xml", dir=str(registry_path.parent)
    )
    os.close(descriptor)
    try:
        tree.write(temporary_name, encoding="utf-8", xml_declaration=True)
        os.replace(temporary_name, str(registry_path))
    finally:
        if os.path.exists(temporary_name):
            os.unlink(temporary_name)
    return True


def install():
    platform_name = _platform_name()
    architecture = _architecture()
    addons_directory = _addons_directory(platform_name)
    addon_path = addons_directory / ADDIN_DIRECTORY_NAME
    addons_directory.mkdir(parents=True, exist_ok=True)

    registry_path = addons_directory / "publish.xml"
    _load_registry(registry_path)
    auth_token = _credential(platform_name)
    source_digest = _source_digest()
    runtime_config = _runtime_config_content(auth_token)
    existed = addon_path.exists()
    addon_changed = not _addon_is_current(
        addon_path, source_digest, runtime_config
    )
    if addon_changed:
        _replace_addon(
            addons_directory, addon_path, source_digest, runtime_config
        )
    registration_changed = _merge_registration(registry_path)
    changed = addon_changed or registration_changed
    return {
        "platform": platform_name,
        "architecture": architecture,
        "status": "updated" if existed and changed else "installed" if changed else "current",
        "restart_required": changed,
        "addon_path": str(addon_path),
    }
