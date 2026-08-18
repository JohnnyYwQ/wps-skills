# WPS Office Skill installation

The `wps-office` directory is the complete distributable package. It requires Python 3.9+ and WPS Office; the Runner and Add-in resources are already included.

## Install and check

Copy or link `skills/wps-office/` into the host's skill directory, then run the following from that directory:

```bash
python3 scripts/wps.py check
```

On first use, `check` installs the Add-in in the current user's WPS profile and writes a local credential. If the response status is `restart_required`, fully exit and reopen WPS Office, then run `check` again. Continue only when `data.ready` is `true`.

## Invoke an Action

Read `SKILL.md`, `references/action-manifest.json`, and the applicable Excel, Word, or PowerPoint reference before invoking an Action:

```bash
python3 scripts/wps.py invoke '{"action":"ping","params":{},"timeout_ms":30000}'
```

Each call starts a temporary loopback service, waits for the installed Add-in to execute one WPS Action, prints one JSON result to standard output, and exits.

For a destructive Action, explain the exact consequence and obtain explicit user confirmation before including `"confirmed": true` in the request. Without it, the Runner returns `CONFIRMATION_REQUIRED` before contacting WPS.

## Troubleshooting

- `wps_not_running`: start WPS Office, then run `check` again.
- `addin_unavailable`: fully restart WPS Office, then rerun `check`.
- `ACTION_BUSY`: wait for the current call to finish and retry.
- `PORT_IN_USE`: release the process using `127.0.0.1:58891`, then retry.

See `SKILL.md` for the complete readiness and error-recovery protocol.
