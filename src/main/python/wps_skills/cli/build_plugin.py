"""Build a relocatable Codex/Claude Code plugin distribution from production Skills."""

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import tempfile
import zipfile

from wps_skills.cli.build_skill import MAIN, build_application_skill


PLUGIN_NAME = "wps-skills"
MARKETPLACE_NAME = "wps-skills-local"
PLUGIN_SOURCE = MAIN / "resources" / "plugins" / PLUGIN_NAME
REPOSITORY = MAIN.parents[1]
APPLICATIONS = ("word", "excel", "ppt")


def write_json(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def build_plugin(destination):
    """Return a new distribution directory; never overwrite an earlier build."""
    destination = Path(destination).resolve()
    archive_path = destination.parent / (destination.name + ".zip")
    checksum_path = destination.parent / (destination.name + ".zip.sha256")
    for path in (destination, archive_path, checksum_path):
        if path.exists():
            raise FileExistsError("Build output already exists: " + str(path))
    manifest = json.loads((PLUGIN_SOURCE / ".codex-plugin/plugin.json").read_text(encoding="utf-8"))
    if manifest["name"] != PLUGIN_NAME:
        raise ValueError("Plugin name must match its directory: " + PLUGIN_NAME)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="wps-plugin-build-", dir=destination.parent) as temporary:
        stage = Path(temporary) / destination.name
        plugin = stage / "plugins" / PLUGIN_NAME
        write_json(plugin / ".codex-plugin/plugin.json", manifest)
        # Claude discovers skills/ by convention. Keep host-specific UI metadata
        # in the Codex manifest, with a single source for shared identity/version.
        claude = {key: manifest[key] for key in (
            "name", "version", "description", "author", "homepage", "repository", "license", "keywords")}
        write_json(plugin / ".claude-plugin/plugin.json", claude)
        for app in APPLICATIONS:
            build_application_skill(app, plugin / "skills" / ("wps-" + app))
        relative_plugin = "./plugins/" + PLUGIN_NAME
        write_json(stage / ".agents/plugins/marketplace.json", {
            "name": MARKETPLACE_NAME,
            "interface": {"displayName": "WPS Skills"},
            "plugins": [{
                "name": PLUGIN_NAME,
                "source": {"source": "local", "path": relative_plugin},
                "policy": {"installation": "AVAILABLE", "authentication": "ON_INSTALL"},
                "category": "Productivity",
            }],
        })
        write_json(stage / ".claude-plugin/marketplace.json", {
            "name": MARKETPLACE_NAME,
            "owner": manifest["author"],
            "description": manifest["description"],
            "plugins": [{"name": PLUGIN_NAME, "source": relative_plugin}],
        })
        shutil.copy2(PLUGIN_SOURCE / "INSTALL.md", stage / "README.md")
        shutil.copy2(PLUGIN_SOURCE / "USAGE.md", plugin / "README.md")
        for root in (stage, plugin):
            shutil.copy2(REPOSITORY / "LICENSE", root / "LICENSE")
        files = {path.relative_to(stage).as_posix(): hashlib.sha256(path.read_bytes()).hexdigest()
                 for path in sorted(stage.rglob("*")) if path.is_file()}
        write_json(stage / "files.sha256.json", {
            "plugin": PLUGIN_NAME, "version": manifest["version"], "files": files,
        })
        temporary_archive = Path(temporary) / "distribution.zip"
        with zipfile.ZipFile(temporary_archive, "x", zipfile.ZIP_DEFLATED) as archive:
            for path in sorted(stage.rglob("*")):
                if path.is_file():
                    archive.write(path, (Path(destination.name) / path.relative_to(stage)).as_posix())
        checksum = hashlib.sha256(temporary_archive.read_bytes()).hexdigest()
        # Publish only after every Skill and archive has been assembled.
        stage.rename(destination)
        with archive_path.open("xb") as output, temporary_archive.open("rb") as source:
            shutil.copyfileobj(source, output)
        with checksum_path.open("x", encoding="utf-8") as output:
            output.write(checksum + "  " + archive_path.name + "\n")
    return destination


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=Path("build/plugins/wps-skills"),
                        help="New distribution directory; a ZIP and SHA-256 sidecar are also generated")
    args = parser.parse_args(argv)
    try:
        destination = build_plugin(args.output)
    except (OSError, ValueError) as error:
        parser.exit(2, str(error) + "\n")
    print(destination)
    print(destination.parent / (destination.name + ".zip"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
