#!/usr/bin/env python3
"""Mac entry for isolated Windows startup/communication experiments (milestone 1)."""
import argparse
import base64
from datetime import datetime, timezone
import json
from pathlib import Path
import re
import shutil
import statistics
import subprocess
import sys
import uuid
import zipfile

from bundle import HERE, REPO, RESOURCES, build
from common import Journal, digest, read_json, utc, validate_config, write_json


def ps_literal(value):
    return "'" + str(value).replace("'", "''") + "'"


def remote(host, script, directory, name, timeout=90):
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.@-]*", host):
        raise ValueError("Use an SSH alias or simple user@host")
    script = "$ErrorActionPreference='Stop'; $ProgressPreference='SilentlyContinue'; [Console]::OutputEncoding=New-Object Text.UTF8Encoding($false);\n" + script
    # Windows OpenSSH may execute through cmd.exe (8191-character command limit).
    # Keep the bootstrap short; send the actual script over UTF-8 stdin.
    bootstrap = "[Console]::InputEncoding=New-Object Text.UTF8Encoding($false); $body=[Console]::In.ReadToEnd(); & ([scriptblock]::Create($body))"
    encoded = base64.b64encode(bootstrap.encode("utf-16-le")).decode("ascii")
    return command(["ssh", "-o", "BatchMode=yes", "-o", "ConnectTimeout=10", host,
                    "powershell.exe -NoLogo -NoProfile -NonInteractive -EncodedCommand " + encoded], directory, name, timeout, input_text=script)


def command(argv, directory, name, timeout=90, input_text=None):
    journal = Journal(directory / "host-events.jsonl", clockDomain="mac", runId=directory.name)
    log = directory / "host-logs" / (name + "-" + uuid.uuid4().hex[:8])
    log.parent.mkdir(exist_ok=True)
    with journal.phase(name):
        try:
            cp = subprocess.run(argv, input=input_text, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout)
        except subprocess.TimeoutExpired as error:
            Path(str(log) + ".error.txt").write_text(str(error), encoding="utf-8")
            raise RuntimeError(name + " timed out; inspect status before taking any further action") from error
        Path(str(log) + ".stdout.txt").write_text(cp.stdout, encoding="utf-8")
        Path(str(log) + ".stderr.txt").write_text(cp.stderr, encoding="utf-8")
        if cp.returncode:
            raise RuntimeError(f"{name} failed ({cp.returncode}); see {log}.stderr.txt")
        return cp.stdout.strip("\ufeff\r\n ")


def pack(args):
    run_id = "startup-" + datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S") + "-" + uuid.uuid4().hex[:8]
    directory = args.output.resolve() / run_id
    directory.mkdir(parents=True, exist_ok=False)
    journal = Journal(directory / "host-events.jsonl", clockDomain="mac", runId=run_id)
    with journal.phase("workspace.package"):
        archive = build(directory / "bundle", args.config)
    write_json(directory / "run.json", {"runId": run_id, "state": "packed", "createdUtc": utc(),
               "archiveSha256": digest(archive), "profile": args.profile})
    print(str(directory))
    return directory


def start(args):
    directory = args.run.resolve()
    metadata = read_json(directory / "run.json")
    if metadata["state"] != "packed":
        raise ValueError("This run was already submitted or submission is uncertain. Use status; pack a new run for new execution.")
    if digest(directory / "bundle.zip") != metadata["archiveSha256"]:
        raise ValueError("Archive changed after packing")
    config = validate_config(read_json(directory / "bundle/experiment.json"))
    run_id = metadata["runId"]
    if not re.fullmatch(r"startup-[A-Za-z0-9-]+", run_id):
        raise ValueError("Invalid run identity")
    relative = "wps-startup-communication/" + run_id
    metadata.update(state="submission_started", host=args.host, remoteRelative=relative)
    write_json(directory / "run.json", metadata)  # before network I/O, never auto-retry
    supplied_python = ps_literal(args.python) if args.python else "$null"
    preparation = f"""
$root = Join-Path $env:USERPROFILE {ps_literal(relative)}
if (Test-Path -LiteralPath $root) {{ throw 'Run directory exists; never repeat submission' }}
$python = {supplied_python}
if (-not $python) {{
    $candidates = @(Get-ChildItem -Path (Join-Path $env:USERPROFILE '.workbuddy/binaries/python/versions/*/python.exe') -ErrorAction SilentlyContinue | Sort-Object FullName -Descending)
    if ($candidates.Count) {{ $python = $candidates[0].FullName }}
    else {{ $python = (Get-Command python.exe -ErrorAction Stop).Source }}
}}
if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {{ throw 'Python executable not found; supply --python' }}
[void](New-Item -ItemType Directory -Path $root)
@{{root=$root; python=$python}} | ConvertTo-Json -Compress
"""
    prepared = json.loads(remote(args.host, preparation, directory, "remote.prepare"))
    metadata.update(remoteRoot=prepared["root"], python=prepared["python"])
    write_json(directory / "run.json", metadata)
    command(["scp", "-o", "BatchMode=yes", "-o", "ConnectTimeout=10", str(directory / "bundle.zip"),
             args.host + ":" + relative + "/bundle.zip"], directory, "remote.upload")
    task_name = "WpsStartup-" + run_id
    metadata.update(taskName=task_name, state="trigger_pending")
    write_json(directory / "run.json", metadata)
    budget = config["profiles"][metadata["profile"]]["budgetSeconds"]
    deadline = int(budget + config["limits"]["caseSeconds"] + 120)
    deployment = f"""
$root = {ps_literal(prepared['root'])}
$archive = Join-Path $root 'bundle.zip'
if ((Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant() -ne {ps_literal(metadata['archiveSha256'])}) {{ throw 'Uploaded archive hash mismatch' }}
Expand-Archive -LiteralPath $archive -DestinationPath (Join-Path $root 'bundle')
$python = {ps_literal(prepared['python'])}
$launch = Join-Path $root 'bundle/launch.ps1'
$args = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $launch + '" -Root "' + $root + '" -Python "' + $python + '" -Profile {metadata['profile']}'
$action = New-ScheduledTaskAction -Execute (Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe') -Argument $args
$principal = New-ScheduledTaskPrincipal -UserId ([Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Seconds {deadline}) -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$name = {ps_literal(task_name)}
if (Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue) {{ throw 'Task exists; will not restart it' }}
Register-ScheduledTask -TaskName $name -Action $action -Principal $principal -Settings $settings | Out-Null
Start-ScheduledTask -TaskName $name
@{{submitted=$true; task=$name; utc=[DateTime]::UtcNow.ToString('o')}} | ConvertTo-Json -Compress
"""
    receipt = json.loads(remote(args.host, deployment, directory, "remote.deploy_and_trigger"))
    metadata.update(state="submitted", submittedUtc=receipt["utc"])
    write_json(directory / "run.json", metadata)
    print("Submitted " + run_id + "; Windows runs independently. Use status --run " + str(directory))


def remote_context(directory):
    metadata = read_json(directory / "run.json")
    if not metadata.get("remoteRoot"):
        raise ValueError("No confirmed remote location; inspect host logs. Do not repeat start.")
    return metadata, "$root = " + ps_literal(metadata["remoteRoot"]) + "\n"


def status(args):
    directory = args.run.resolve()
    metadata, prefix = remote_context(directory)
    script = (RESOURCES / "snapshot_io.ps1").read_text(encoding="utf-8") + "\n" + prefix + f"""
$state = $null; $launch = $null; $tail = @()
$path = Join-Path $root 'windows/status.json'
if (Test-Path -LiteralPath $path) {{ $state = Read-DiagnosticText -Path $path | ConvertFrom-Json }}
$path = Join-Path $root 'windows/launcher.json'
if (Test-Path -LiteralPath $path) {{ $launch = Read-DiagnosticText -Path $path | ConvertFrom-Json }}
$task = Get-ScheduledTask -TaskName {ps_literal(metadata.get('taskName', 'not-registered'))} -ErrorAction SilentlyContinue
$taskInfo = if ($task) {{ Get-ScheduledTaskInfo -InputObject $task }} else {{ $null }}
if ($state -and $state.activeCase) {{
    $events = Join-Path $root ('windows/cases/' + $state.activeCase + '/events.jsonl')
    if (Test-Path -LiteralPath $events) {{
        $tail = @((Read-DiagnosticText -Path $events) -split "`r?`n" | Where-Object {{ $_ }} | Select-Object -Last 3 | ForEach-Object {{
            try {{ ConvertFrom-Json -InputObject ([string]$_) }} catch {{ @{{partial=$true}} }}
        }})
    }}
}}
@{{status=$state; launcher=$launch; scheduledTaskState=[string]$task.State; lastTaskResult=$taskInfo.LastTaskResult; recentEvents=$tail; stopRequested=(Test-Path (Join-Path $root 'stop.request'))}} | ConvertTo-Json -Depth 30 -Compress
"""
    result = json.loads(remote(metadata["host"], script, directory, "remote.status"))
    state = result.get("status")
    if state and state["state"] == "running" and result["scheduledTaskState"] != "Running":
        result["observation"] = "Supervisor is not running but no final status exists: interrupted/unknown; never replay."
    write_json(directory / "last-status.json", result)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    elif state:
        passed = state.get('observedPassed', sum(c['state']=='passed' for c in state['cases']))
        failed = state.get('observedFailed', sum(c['state']=='failed' for c in state['cases']))
        print(f"{metadata['runId']}: {state['state']} | passed={passed} failed={failed} | active={state['activeCase'] or '-'} | {state.get('stopReason') or ''}")
        if state.get("error"):
            print(state["error"])
        if result.get("observation"):
            print(result["observation"])
        if result["recentEvents"]:
            print("Latest: " + json.dumps(result["recentEvents"][-1], ensure_ascii=False))
    else:
        print("No supervisor receipt yet: " + json.dumps(result, ensure_ascii=False))
    return result


def stop(args):
    directory = args.run.resolve()
    metadata, prefix = remote_context(directory)
    remote(metadata["host"], prefix + "[IO.File]::WriteAllText((Join-Path $root 'stop.request'), [DateTime]::UtcNow.ToString('o'))", directory, "remote.request_stop")
    print("Stop requested. Current case finishes or hits its watchdog; no next case starts. Use status to verify.")


def extract_snapshot(archive, target):
    with zipfile.ZipFile(archive) as source:
        for info in source.infolist():
            normalized = info.filename.replace("\\", "/")
            candidate = (target / normalized).resolve()
            try:
                candidate.relative_to(target.resolve())
            except ValueError:
                raise ValueError("Unsafe snapshot entry")
        for info in source.infolist():
            normalized = info.filename.replace("\\", "/")
            candidate = target / normalized
            if normalized.endswith("/"):
                candidate.mkdir(parents=True, exist_ok=True)
            else:
                candidate.parent.mkdir(parents=True, exist_ok=True)
                with source.open(info) as incoming, candidate.open("wb") as outgoing:
                    shutil.copyfileobj(incoming, outgoing)


def collect(args):
    directory = args.run.resolve()
    metadata, prefix = remote_context(directory)
    snapshot_id = "snapshot-" + uuid.uuid4().hex
    # Snapshot is materialized remotely before transfer. Running snapshots are explicitly partial.
    script = (RESOURCES / "snapshot_io.ps1").read_text(encoding="utf-8") + "\n" + prefix + f"""
Add-Type -AssemblyName System.IO.Compression.FileSystem
$snapshot = Join-Path $root {ps_literal(snapshot_id)}
[void](New-Item -ItemType Directory -Path $snapshot)
$omitted = @()
$files = @(Get-ChildItem -LiteralPath (Join-Path $root 'windows') -Recurse -File -ErrorAction SilentlyContinue)
foreach ($name in @('evidence.sha256.json','stop.request')) {{ if (Test-Path (Join-Path $root $name)) {{ $files += Get-Item -LiteralPath (Join-Path $root $name) }} }}
foreach ($file in $files) {{
    $relative = $file.FullName.Substring($root.Length).TrimStart([char[]]@([char]92, [char]47))
    $target = Join-Path $snapshot $relative
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target))
    try {{ Copy-DiagnosticEvidenceFile -Source $file.FullName -Destination $target }}
    catch {{ $omitted += @{{file=$relative; error=$_.Exception.Message}} }}
}}
$state = $null
if (Test-Path (Join-Path $snapshot 'windows/status.json')) {{ $state = Read-DiagnosticText -Path (Join-Path $snapshot 'windows/status.json') | ConvertFrom-Json }}
$complete = ($null -ne $state -and $state.state -ne 'running' -and $omitted.Count -eq 0 -and (Test-Path (Join-Path $snapshot 'evidence.sha256.json')))
[IO.File]::WriteAllText((Join-Path $snapshot 'snapshot.json'), (@{{capturedUtc=[DateTime]::UtcNow.ToString('o'); finalized=$complete; omitted=$omitted}} | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))
$archive = $snapshot + '.zip'
[IO.Compression.ZipFile]::CreateFromDirectory($snapshot, $archive)
@{{sha256=(Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant()}} | ConvertTo-Json -Compress
"""
    receipt = json.loads(remote(metadata["host"], script, directory, "remote.snapshot"))
    destination = directory / "snapshots" / snapshot_id
    destination.mkdir(parents=True, exist_ok=False)
    archive = destination / "evidence.zip"
    command(["scp", "-o", "BatchMode=yes", "-o", "ConnectTimeout=10",
             metadata["host"] + ":" + metadata["remoteRelative"] + "/" + snapshot_id + ".zip", str(archive)], directory, "remote.download")
    if digest(archive) != receipt["sha256"]:
        raise ValueError("Downloaded snapshot hash mismatch")
    extract_snapshot(archive, destination)
    capture = read_json(destination / "snapshot.json")
    if capture["finalized"]:
        for name, expected in read_json(destination / "evidence.sha256.json").items():
            if digest(destination / "windows" / name) != expected:
                raise ValueError("Evidence hash mismatch: " + name)
    write_json(directory / "latest-snapshot.json", {"path": destination.relative_to(directory).as_posix(), "finalized": capture["finalized"]})
    print(str(destination) + (" (finalized, hashes verified)" if capture["finalized"] else " (partial snapshot)"))


def snapshot_directory(directory, snapshot):
    # Keep historical absolute pointers intact; anchor their snapshot identity
    # within the current run so archives remain movable and self-contained.
    name = snapshot["path"].replace("\\", "/").rsplit("/", 1)[-1]
    if not re.fullmatch(r"snapshot-[0-9a-f]{32}", name):
        raise ValueError("Invalid snapshot identity")
    return directory / "snapshots" / name


def report(args):
    directory = args.run.resolve()
    snapshot = read_json(directory / "latest-snapshot.json")
    evidence = snapshot_directory(directory, snapshot) / "windows"
    state = read_json(evidence / "status.json")
    print(f"{state['runId']} | {state['state']} | finalized={snapshot['finalized']}")
    print("Case                         State      Seconds")
    history = evidence / state.get('caseHistory', 'cases.jsonl')
    cases = [json.loads(line) for line in history.read_text(encoding='utf-8').splitlines() if line] if history.exists() else state['cases']
    for case in cases:
        print(f"{case['name']:28} {case['state']:10} {case.get('durationMs', 0)/1000:.3f}")
        if case["state"] == "failed":
            print("  " + json.dumps(case.get("result"), ensure_ascii=False))
    from com_report import print_summary
    for path in sorted(evidence.glob("cases/*/com-summary.json")):
        print_summary(read_json(path), path.parent.name)
    samples = {}
    for path in evidence.glob("cases/*/events.jsonl"):
        for line in path.read_text(encoding="utf-8").splitlines():
            try:
                event = json.loads(line)
            except json.JSONDecodeError:
                continue  # a partial snapshot may end inside a record
            if event.get("event") == "phase.finished" and event.get("outcome") == "succeeded":
                key = (path.parent.name.split("-", 1)[1], event["name"], event.get("operation", ""))
                samples.setdefault(key, []).append(event["durationMs"])
    print("\nMeasured phase (nested; do not sum)                         N  median ms     max ms")
    for key, values in sorted(samples.items()):
        print(f"{' / '.join(filter(None, key)):58} {len(values):3} {statistics.median(values):10.3f} {max(values):10.3f}")
    print("\nTest orchestration (separate from product execution)")
    for label, path in (("Mac", directory / "host-events.jsonl"), ("Windows", evidence / "events.jsonl")):
        stages = {}
        if path.exists():
            for line in path.read_text(encoding="utf-8").splitlines():
                try:
                    event = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if event.get("event") == "phase.finished":
                    stages.setdefault((event["name"], event["outcome"]), []).append(event["durationMs"])
        for (name, outcome), values in sorted(stages.items()):
            print(f"{label:8} {name:32} {outcome:10} n={len(values)} median={statistics.median(values):.3f} ms max={max(values):.3f} ms")
    if args.detail:
        print("\nProduction trace phases (nested; wait_response includes native execution)")
        phases = {}
        for path in evidence.glob("cases/*/traces/**/*.jsonl"):
            for line in path.read_text(encoding="utf-8-sig").splitlines():
                try:
                    event = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if event.get("event") in ("span.finished", "phase.finished") and "durationMs" in event:
                    key = (path.parts[path.parts.index("cases") + 1].split("-", 1)[1], event["name"])
                    phases.setdefault(key, []).append(event["durationMs"])
        for key, values in sorted(phases.items()):
            print(f"{' / '.join(key):58} {len(values):3} {statistics.median(values):10.3f} {max(values):10.3f}")
    print("\nObserved baseline only. No alternative-scheme comparison or reliability guarantee.")
    print("Visual review: " + state.get("visualReview", "not_yet_evaluated"))
    print("Evidence: " + str(evidence))


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    p = commands.add_parser("pack", help="Snapshot current workspace; no network or execution")
    p.add_argument("--config", type=Path, default=RESOURCES / "quick.json")
    p.add_argument("--profile", choices=("quick", "long"), default="quick")
    p.add_argument("--output", type=Path, default=REPO / "build/runs/startup-communication")
    p.set_defaults(function=pack)
    p = commands.add_parser("start", help="Upload and trigger exactly once; returns immediately")
    p.add_argument("--run", required=True, type=Path)
    p.add_argument("--host", default="win")
    p.add_argument("--python", help="Absolute Windows Python executable; auto-detected when omitted")
    p.set_defaults(function=start)
    for name, function in (("status", status), ("collect", collect), ("stop", stop), ("report", report)):
        p = commands.add_parser(name)
        p.add_argument("--run", required=True, type=Path)
        if name == "status":
            p.add_argument("--json", action="store_true")
        if name == "report":
            p.add_argument("--detail", action="store_true", help="Include existing Task/Action/bridge/native trace phase aggregates")
        p.set_defaults(function=function)
    args = parser.parse_args(argv)
    try:
        args.function(args)
    except (ValueError, OSError, RuntimeError, subprocess.SubprocessError) as error:
        parser.exit(1, str(error) + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
