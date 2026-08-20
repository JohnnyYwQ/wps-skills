---
name: wps-office
description: Operate WPS Office through canonical WPS Actions and the bundled Python 3.9+ Runner. Use for setup, readiness checks, or read and write Excel, Word, PowerPoint, common, and cross-application work through WPS Office.
---

# WPS Office

Execute one WPS Action per Python process. Use only the bundled standard-library Runner; do not start a separate service or use another runtime.

## Check Readiness

Before every Action workflow, run:

```bash
python3 scripts/wps.py check
```

The first check installs or updates the bundled WPS Add-in in the current user's profile on Linux, macOS, or Windows. On macOS it uses WPS's container profile and clears the copied Add-in's Gatekeeper quarantine attribute. Handle `data.status` as follows:

- `ready`: continue to the Action.
- `restart_required`: ask the user to fully exit and restart WPS Office, then check again.
- `wps_not_running`: ask the user to start WPS Office, then check again.
- `addin_unavailable`: ask the user to fully restart WPS Office and check again; if it persists, report that the installed Add-in is not responding.

When a non-ready check includes `data.error`, use its stable transport code to explain the cause; the readiness `status` remains the recovery state for compatibility.

Do not run the Action unless `data.ready` is `true`. The installer recognizes Linux x86_64, Linux ARM64, macOS x86_64, macOS ARM64, Windows x86_64, and Windows ARM64. macOS installation is experimental and outside the supported-target matrix until its matching real-WPS acceptance suite passes. Setup failures return a stable structured error on stdout and a nonzero exit status.

## Choose an Action

Read `references/action-manifest.json` and select the canonical WPS Action matching the request. Check its `application`, parameter contract, result contract, prerequisites, and `risk` before invoking it.

For opening, saving, saving as, converting, or closing files across applications, also read `references/common.md`. For Excel core data, formatting, chart, analysis, pivot, protection, image, or context-aware work, also read `references/excel.md`. For Word document lifecycle, content, formatting, template, bookmark, comment, or revision work, also read `references/word.md`. For PowerPoint presentation lifecycle, slide, text, image, table, note, background, basic layout, advanced design, chart, diagram, animation, transition, theme, master, or cross-presentation work, also read `references/powerpoint.md`. Use the relevant workflow guidance while treating the manifest as the exact parameter and result contract.

Apply the manifest risk policy before invoking:

- `read`: invoke directly after readiness succeeds.
- `write`: invoke when the user's current request authorizes the change; no confirmation marker is required.
- `destructive`: explain the specific deletion, overwrite, or discard and obtain explicit user confirmation. Only after that confirmation, invoke with `"confirmed":true`. Never infer confirmation from an earlier general request.

Do not invoke an Action whose `risk` is missing or is not one of these three values. The Runner also rejects that invalid manifest state before contacting WPS.

## Invoke the Runner

After readiness succeeds, run from this skill directory. Pass exactly one JSON request as the final argument:

```bash
python3 scripts/wps.py invoke '{"action":"ping","params":{},"timeout_ms":30000}'
```

Replace `ping` and `params` with the chosen Action. Keep `timeout_ms` bounded. Each invocation starts one temporary service bound to `127.0.0.1:58891`, waits for the WPS Add-in to poll and return the correlated result, then closes the service and exits.

For a destructive Action only, include the confirmation marker after the user explicitly confirms the named consequence:

```bash
python3 scripts/wps.py invoke '{"action":"deleteSlide","params":{"slideIndex":2},"confirmed":true,"timeout_ms":30000}'
```

If confirmation is absent, the Runner returns `CONFIRMATION_REQUIRED` before the WPS Add-in receives the Action. Do not retry until explicit confirmation is obtained. `INVALID_ACTION_RISK` means the manifest risk is missing or unknown; do not bypass or retry that gate.

Readiness checks and invocations share a per-user cross-process lock. If another WPS Action owns it, `ACTION_BUSY` is retryable after that process finishes. Never start a background service to work around this error.

Parse stdout as one JSON object. A success has this shape:

```json
{"ok":true,"action":"ping","data":{"message":"pong","timestamp":1723852800000}}
```

A failure returns a nonzero exit code and a structured stdout result:

```json
{"ok":false,"action":"ping","error":{"code":"ADDIN_NOT_READY","message":"WPS Add-in did not return a result before the timeout","retryable":true}}
```

Retry only when `error.retryable` is `true`. Report non-retryable WPS Action failures to the user without changing their message.

Use transport error codes to explain recovery precisely:

- `ADDIN_NOT_READY`: no authenticated WPS Add-in polled before the deadline; check readiness again.
- `ACTION_TIMEOUT`: WPS accepted the Action but did not finish before the deadline; retry only when the returned metadata permits it.
- `PORT_IN_USE`: another process owns the loopback port; stop that conflicting process before retrying.
- `PORT_UNAVAILABLE`: the Runner could not open its loopback port; correct the local socket or permission problem before retrying.
- `ADDIN_DISCONNECTED`, `INVALID_ADDIN_JSON`, `INVALID_ADDIN_RESPONSE`, or `REQUEST_ID_MISMATCH`: the authenticated WPS Add-in failed the result protocol; do not retry automatically.

The Runner releases its loopback port and Action lock on every success or failure path. Do not print, copy, or log the installed authentication credential or Action parameters while diagnosing transport failures.
