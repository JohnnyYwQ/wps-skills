"""Assemble a standalone Application Skill without maintaining a second runtime copy."""

import hashlib
import json
from pathlib import Path
import shutil
import tempfile


MAIN = Path(__file__).resolve().parents[3]


def write_action_definitions(application, skill_directory, *, combined_export=True):
    """Export the production contracts as a name-keyed, relocatable reference."""
    import importlib
    from wps_skills.client.applications import contracts_for
    definitions = {contract.name: contract.to_wire(application)
                   for contract in contracts_for(application).contracts}
    schemas_module = importlib.import_module("wps_skills.word.contracts.schemas" if application == "word"
                                            else "wps_skills." + application + ".contracts")
    destination = Path(skill_directory) / "references" / "actions.json"
    if combined_export:
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(
            json.dumps(definitions, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
    from wps_skills.cli.schema_references import write_schema_references
    schemas = write_schema_references(skill_directory, definitions, schemas=schemas_module)
    return destination if combined_export else schemas


def build_application_skill(application, destination):
    from wps_skills.client.applications import profile
    profile(application)
    skill_name = "wps-" + application
    destination = Path(destination).resolve()
    if destination.exists():
        raise FileExistsError(f"Destination already exists: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=skill_name + "-build-", dir=destination.parent) as temporary:
        stage = Path(temporary) / skill_name
        skill_source = MAIN / "resources" / "skills" / skill_name
        # Publish only the entry points and host metadata. References are
        # generated from admitted contracts, never copied from a workspace
        # that can contain old exports, experiments or development notes.
        for relative in ("SKILL.md", "agents/openai.yaml",
                         "scripts/" + application + ".py", "scripts/schema.py"):
            target = stage / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            source = MAIN / "resources/skills/wps-word/scripts/schema.py" if relative == "scripts/schema.py" else skill_source / relative
            shutil.copy2(source, target)
        write_action_definitions(application, stage, combined_export=False)
        runtime = stage / "runtime" / "src" / "main"
        python_source = MAIN / "python" / "wps_skills"
        python_target = runtime / "python" / "wps_skills"
        python_target.mkdir(parents=True)
        shutil.copy2(python_source / "__init__.py", python_target / "__init__.py")
        # Each independently installed Skill carries its own shared runtime copy.
        # Other applications and repository build/demo tools are not dependencies.
        runtime_ignore = shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store", "demo.py", "desktop.py")
        packages = ("core", "client", "windows", application)
        for package in packages:
            shutil.copytree(python_source / package, python_target / package, ignore=runtime_ignore)
        (python_target / "cli").mkdir()
        cli_files = ("__init__.py", "discovery.py", "task_io.py", "task.py")
        for name in cli_files:
            shutil.copy2(python_source / "cli" / name, python_target / "cli" / name)
        resource_source = MAIN / "resources" / "wps_skills"
        resource_target = runtime / "resources" / "wps_skills"
        resource_ignore = shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store", "demo_launcher.ps1", "demo-template.pptx")
        for package in ("windows", application):
            shutil.copytree(resource_source / package, resource_target / package, ignore=resource_ignore)
        files = {
            path.relative_to(stage).as_posix(): hashlib.sha256(path.read_bytes()).hexdigest()
            for path in sorted(stage.rglob("*")) if path.is_file()
        }
        (stage / "runtime" / "files.sha256.json").write_text(
            json.dumps(files, indent=2) + "\n", encoding="utf-8"
        )
        stage.rename(destination)
    return destination


def write_word_action_definitions(skill_directory, *, combined_export=True):
    return write_action_definitions("word", skill_directory, combined_export=combined_export)
