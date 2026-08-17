---
name: wps-office
description: Operate WPS Office through canonical WPS Actions and the bundled Python 3.9+ Runner. Use for read-only Excel, Word, PowerPoint, common, or cross-application inspection through WPS Office.
---

# WPS Office

Execute one read-only WPS Action per Python process. Use only the bundled standard-library Runner; do not start a separate service or use another runtime.

## Choose an Action

Read `references/action-manifest.json` and select the canonical WPS Action matching the request. Check its `application`, parameter contract, result contract, prerequisites, and `risk` before invoking it.

This initial package supports the end-to-end read path. Invoke only entries whose `risk` is `read`. Do not invoke `write` or `destructive` entries until their Runner safety gates are available.

## Invoke the Runner

Run from this skill directory. Pass exactly one JSON request as the final argument:

```bash
python3 scripts/wps.py invoke '{"action":"ping","params":{},"timeout_ms":30000}'
```

Replace `ping` and `params` with the chosen read Action. Keep `timeout_ms` bounded. Each invocation starts one temporary service bound to `127.0.0.1:58891`, waits for the WPS Add-in to poll and return the correlated result, then closes the service and exits.

Parse stdout as one JSON object. A success has this shape:

```json
{"ok":true,"action":"ping","data":{"message":"pong","timestamp":1723852800000}}
```

A failure returns a nonzero exit code and a structured stdout result:

```json
{"ok":false,"action":"ping","error":{"code":"ADDIN_NOT_READY","message":"WPS Add-in did not return a result before the timeout","retryable":true}}
```

Retry only when `error.retryable` is `true`. Report non-retryable WPS Action failures to the user without changing their message.
