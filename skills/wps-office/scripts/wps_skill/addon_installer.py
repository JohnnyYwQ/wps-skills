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
TRANSACTION_MARKER = ".wps-office-skill-transaction.json"
ADDIN_BACKUP = ".wps-office-skill-backup"
REGISTRY_BACKUP = ".wps-office-skill-publish-backup.xml"


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


def action_lock_path():
    """Return the current user's cross-process WPS Action lock path."""
    return _config_path(_platform_name()).parent / "action.lock"


def source_digest():
    """Return the digest the loaded Add-in must acknowledge."""
    return _source_digest()


def restart_is_pending(platform_name):
    config_path = _config_path(platform_name)
    try:
        config = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, TypeError, ValueError):
        return True
    return config.get("loaded_digest") != _source_digest()


def acknowledge_loaded_digest(platform_name, loaded_digest):
    expected_digest = _source_digest()
    if loaded_digest != expected_digest:
        return False
    config_path = _config_path(platform_name)
    try:
        config = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, TypeError, ValueError) as error:
        raise AddinInstallError(
            "INVALID_AUTH_CONFIG",
            "The WPS Skill authentication config is invalid: {0}".format(error),
        )
    if config.get("loaded_digest") == loaded_digest:
        return True
    config["loaded_digest"] = loaded_digest
    _write_private_text(
        config_path,
        json.dumps(config, ensure_ascii=False, separators=(",", ":")) + "\n",
    )
    return True


def wps_is_running(platform_name):
    test_value = os.environ.get("WPS_SKILL_TEST_WPS_RUNNING")
    if test_value is not None:
        return test_value == "1"
    process_names = {
        "wps",
        "wps.exe",
        "et",
        "et.exe",
        "wpp",
        "wpp.exe",
        "wpsoffice",
        "wpsoffice.exe",
    }
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


def _tree_digest(root, excluded_names=()):
    digest = hashlib.sha256()
    source_paths = sorted(
        path
        for path in root.rglob("*")
        if path.is_file() and path.name not in excluded_names
    )
    for source_path in source_paths:
        relative_path = source_path.relative_to(root).as_posix()
        digest.update(relative_path.encode("utf-8"))
        digest.update(b"\0")
        with source_path.open("rb") as source_file:
            for chunk in iter(lambda: source_file.read(65536), b""):
                digest.update(chunk)
    return digest.hexdigest()


def _source_digest():
    return _tree_digest(ADDIN_SOURCE)


def _runtime_config_content(auth_token, source_digest):
    return (
        "var WPS_SKILL_AUTH_TOKEN = {0};\n"
        "var WPS_SKILL_INSTALL_DIGEST = {1};\n"
    ).format(json.dumps(auth_token), json.dumps(source_digest))


def _installed_metadata(addon_path):
    try:
        with (addon_path / INSTALL_METADATA).open(encoding="utf-8") as metadata_file:
            metadata = json.load(metadata_file)
        return metadata
    except (OSError, TypeError, ValueError):
        return {}


def _addon_is_current(addon_path, source_digest, runtime_config):
    if not addon_path.is_dir():
        return False
    metadata = _installed_metadata(addon_path)
    if (
        metadata.get("version") != 1
        or metadata.get("source_digest") != source_digest
    ):
        return False
    try:
        installed_digest = _tree_digest(
            addon_path, excluded_names=(INSTALL_METADATA, RUNTIME_CONFIG)
        )
        return (
            installed_digest == source_digest
            and (addon_path / RUNTIME_CONFIG).read_text(encoding="utf-8")
            == runtime_config
        )
    except OSError:
        return False


def _stage_addon(addons_directory, source_digest, runtime_config):
    staging_root = Path(
        tempfile.mkdtemp(
            prefix=".wps-office-skill-stage-", dir=str(addons_directory)
        )
    )
    staged_addon = staging_root / ADDIN_DIRECTORY_NAME
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
        return staging_root, staged_addon
    except BaseException:
        shutil.rmtree(str(staging_root), ignore_errors=True)
        raise


def _remove_path(path):
    if path.is_dir():
        shutil.rmtree(str(path))
    elif path.exists():
        path.unlink()


def _recover_interrupted_install(addons_directory, addon_path, registry_path):
    marker = addons_directory / TRANSACTION_MARKER
    addon_backup = addons_directory / ADDIN_BACKUP
    registry_backup = addons_directory / REGISTRY_BACKUP
    if not marker.exists():
        _remove_path(addon_backup)
        _remove_path(registry_backup)
        return False
    try:
        transaction = json.loads(marker.read_text(encoding="utf-8"))
    except (OSError, TypeError, ValueError) as error:
        raise AddinInstallError(
            "INCOMPLETE_ADDIN_INSTALL",
            "The prior WPS Add-in update could not be recovered: {0}".format(error),
        )
    if transaction.get("phase") == "committed":
        _remove_path(addon_backup)
        _remove_path(registry_backup)
        marker.unlink()
        return True
    if transaction.get("replace_addon"):
        if transaction.get("had_addon") and addon_backup.exists():
            _remove_path(addon_path)
            os.replace(str(addon_backup), str(addon_path))
        elif not transaction.get("had_addon"):
            _remove_path(addon_path)
    if transaction.get("replace_registry"):
        if transaction.get("had_registry") and registry_backup.exists():
            os.replace(str(registry_backup), str(registry_path))
        elif not transaction.get("had_registry"):
            _remove_path(registry_path)
    marker.unlink()
    return False


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


def _merged_registration(registry_path):
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
        return tree, False

    if registrations:
        registration = registrations[0]
        for duplicate in registrations[1:]:
            root.remove(duplicate)
        registration.attrib.update(expected)
    else:
        ElementTree.SubElement(root, "jsplugin", expected)

    return tree, True


def _stage_registry(tree, registry_path):
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=".publish.", suffix=".xml", dir=str(registry_path.parent)
    )
    os.close(descriptor)
    try:
        tree.write(temporary_name, encoding="utf-8", xml_declaration=True)
        return Path(temporary_name)
    except BaseException:
        os.unlink(temporary_name)
        raise


def _atomic_copy_file(source, destination, validate_xml=False):
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=".{0}.".format(destination.name), dir=str(destination.parent)
    )
    os.close(descriptor)
    temporary_path = Path(temporary_name)
    try:
        shutil.copy2(str(source), str(temporary_path))
        if validate_xml:
            ElementTree.parse(str(temporary_path))
        os.replace(str(temporary_path), str(destination))
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def _commit_install(
    addons_directory,
    addon_path,
    staged_addon,
    registry_path,
    staged_registry,
):
    marker = addons_directory / TRANSACTION_MARKER
    addon_backup = addons_directory / ADDIN_BACKUP
    registry_backup = addons_directory / REGISTRY_BACKUP
    had_addon = addon_path.exists()
    had_registry = registry_path.exists()
    _remove_path(addon_backup)
    _remove_path(registry_backup)
    if staged_registry is not None and had_registry:
        _atomic_copy_file(registry_path, registry_backup, validate_xml=True)
    _write_private_text(
        marker,
        json.dumps(
            {
                "had_addon": had_addon,
                "had_registry": had_registry,
                "replace_addon": staged_addon is not None,
                "replace_registry": staged_registry is not None,
                "phase": "prepared",
            },
            separators=(",", ":"),
        )
        + "\n",
    )
    try:
        if staged_addon is not None:
            if had_addon:
                os.replace(str(addon_path), str(addon_backup))
            os.replace(str(staged_addon), str(addon_path))
        if staged_registry is not None:
            os.replace(str(staged_registry), str(registry_path))
            if had_registry:
                _atomic_copy_file(
                    registry_backup,
                    Path(str(registry_path) + ".bak"),
                    validate_xml=True,
                )
        _write_private_text(
            marker,
            json.dumps(
                {
                    "had_addon": had_addon,
                    "had_registry": had_registry,
                    "replace_addon": staged_addon is not None,
                    "replace_registry": staged_registry is not None,
                    "phase": "committed",
                },
                separators=(",", ":"),
            )
            + "\n",
        )
        _remove_path(addon_backup)
        _remove_path(registry_backup)
        marker.unlink()
    except BaseException:
        committed = _recover_interrupted_install(
            addons_directory, addon_path, registry_path
        )
        if not committed:
            raise


def _install():
    platform_name = _platform_name()
    architecture = _architecture()
    addons_directory = _addons_directory(platform_name)
    addon_path = addons_directory / ADDIN_DIRECTORY_NAME
    addons_directory.mkdir(parents=True, exist_ok=True)

    registry_path = addons_directory / "publish.xml"
    _recover_interrupted_install(addons_directory, addon_path, registry_path)
    _load_registry(registry_path)
    auth_token = _credential(platform_name)
    source_digest = _source_digest()
    runtime_config = _runtime_config_content(auth_token, source_digest)
    existed = addon_path.exists()
    addon_changed = not _addon_is_current(
        addon_path, source_digest, runtime_config
    )
    registry_tree, registration_changed = _merged_registration(registry_path)
    staging_root = None
    staged_addon = None
    staged_registry = None
    try:
        if addon_changed:
            staging_root, staged_addon = _stage_addon(
                addons_directory, source_digest, runtime_config
            )
        if registration_changed:
            staged_registry = _stage_registry(registry_tree, registry_path)
        if addon_changed or registration_changed:
            _commit_install(
                addons_directory,
                addon_path,
                staged_addon,
                registry_path,
                staged_registry,
            )
    finally:
        if staging_root is not None:
            shutil.rmtree(str(staging_root), ignore_errors=True)
        if staged_registry is not None and staged_registry.exists():
            staged_registry.unlink()
    changed = addon_changed or registration_changed
    if existed and changed:
        status = "updated"
    elif changed:
        status = "installed"
    else:
        status = "current"
    return {
        "platform": platform_name,
        "architecture": architecture,
        "status": status,
        "restart_required": changed,
        "addon_path": str(addon_path),
    }


def install():
    try:
        return _install()
    except AddinInstallError:
        raise
    except (OSError, shutil.Error) as error:
        raise AddinInstallError(
            "ADDIN_INSTALL_FAILED",
            "The WPS Add-in could not be installed: {0}".format(error),
        )
