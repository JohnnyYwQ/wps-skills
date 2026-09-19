"""Test-only file formats, validation and evidence recording; standard library only."""
from contextlib import contextmanager
from datetime import datetime, timezone
import hashlib
import json
import math
import os
from pathlib import Path
import time
import uuid


def utc():
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def write_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, allow_nan=False, indent=2) + "\n", encoding="utf-8")
    # Evidence publication only: transient Windows readers must not restart a test case.
    # A permanent access failure still stops the supervisor after this bounded wait.
    deadline = time.monotonic() + 2
    while True:
        try:
            os.replace(temporary, path)
            return
        except PermissionError as error:
            if getattr(error, "winerror", None) not in (5, 32, 33) or time.monotonic() >= deadline:
                raise
            time.sleep(0.025)


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def file_hashes(root):
    return {p.relative_to(root).as_posix(): digest(p) for p in sorted(Path(root).rglob("*"))
            if p.is_file() and "__pycache__" not in p.parts and p.suffix != ".pyc"}


def verify_bundle(root):
    manifest = read_json(Path(root) / "manifest.json")
    for name, expected in manifest["files"].items():
        path = Path(root) / name
        if not path.is_file() or digest(path) != expected:
            raise ValueError("Bundle hash mismatch: " + name)
    return manifest


def validate_config(config):
    def exact(value, keys, label):
        if not isinstance(value, dict) or set(value) != set(keys):
            raise ValueError("Invalid fields: " + label)

    def positive(value, label, integer=False):
        if (isinstance(value, bool) or not isinstance(value, (int, float)) or
                not math.isfinite(value) or value <= 0 or (integer and not isinstance(value, int))):
            raise ValueError("Expected positive " + label)

    exact(config, ("schemaVersion", "name", "claim", "cases", "profiles", "limits", "payloadCharacters", "assertions"), "config")
    if type(config["schemaVersion"]) is not int or config["schemaVersion"] != 1:
        raise ValueError("Unsupported schemaVersion")
    for key in ("name", "claim"):
        if not isinstance(config[key], str) or not config[key].strip():
            raise ValueError("Expected nonempty " + key)
    cases = config["cases"]
    if not isinstance(cases, list) or not cases or any(c not in ("pipe", "wps", "wps-open") for c in cases) or len(set(cases)) != len(cases):
        raise ValueError("cases must contain unique pipe/wps/wps-open entries")
    if "wps-open" in cases and ("wps" not in cases or cases.index("wps") > cases.index("wps-open")):
        raise ValueError("wps-open requires a successful wps fixture earlier in the same round")
    exact(config["profiles"], ("quick", "long"), "profiles")
    for profile in config["profiles"].values():
        exact(profile, ("maxRounds", "budgetSeconds"), "profile")
        positive(profile["budgetSeconds"], "budgetSeconds")
        if profile["maxRounds"] is not None:
            positive(profile["maxRounds"], "maxRounds", True)
    exact(config["limits"], ("startupSeconds", "requestSeconds", "caseSeconds"), "limits")
    for key, value in config["limits"].items():
        positive(value, key)
    if config["limits"]["caseSeconds"] <= config["limits"]["startupSeconds"]:
        raise ValueError("caseSeconds must exceed startupSeconds")
    sizes = config["payloadCharacters"]
    if not isinstance(sizes, list) or not sizes:
        raise ValueError("payloadCharacters must be nonempty")
    for size in sizes:
        positive(size, "payloadCharacters", True)
        if size > 1024 * 1024:
            raise ValueError("Probe payload exceeds 1M characters")
    assertions = config["assertions"]
    exact(assertions, ("exactEcho", "taskOutcome", "cleanupOutcome", "artifactCells", "visualReview"), "assertions")
    if assertions["exactEcho"] is not True or assertions["taskOutcome"] != "succeeded" or assertions["cleanupOutcome"] != "succeeded":
        raise ValueError("Baseline requires exact echo, successful Task and cleanup")
    exact(assertions["artifactCells"], ("A1", "B1"), "artifactCells")
    if not isinstance(assertions["artifactCells"]["A1"], str) or not assertions["artifactCells"]["A1"]:
        raise ValueError("A1 must be nonempty text")
    if type(assertions["artifactCells"]["B1"]) is not int:
        raise ValueError("B1 must be an integer")
    if assertions["visualReview"] not in ("not_required", "required"):
        raise ValueError("Invalid visualReview")
    return config


class Journal:
    def __init__(self, path, **fields):
        self.path = Path(path)
        self.fields = fields

    def event(self, event, **fields):
        record = dict(self.fields, **fields, event=event, utc=utc(),
                      monotonicNs=time.perf_counter_ns(), pid=os.getpid())
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.path.open("a", encoding="utf-8") as stream:
            stream.write(json.dumps(record, ensure_ascii=False, allow_nan=False) + "\n")
            stream.flush()
        return record

    @contextmanager
    def phase(self, name, **fields):
        span = uuid.uuid4().hex
        started = time.perf_counter()
        self.event("phase.started", name=name, spanId=span, **fields)
        try:
            yield
        except BaseException as error:
            self.event("phase.finished", name=name, spanId=span,
                       durationMs=(time.perf_counter() - started) * 1000,
                       outcome="failed", error=repr(error), **fields)
            raise
        else:
            self.event("phase.finished", name=name, spanId=span,
                       durationMs=(time.perf_counter() - started) * 1000,
                       outcome="succeeded", **fields)
