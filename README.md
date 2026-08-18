# WPS Office Skill

`skills/wps-office/` is the sole agent-facing WPS Skill Package. It lets an agent operate Excel, Word, PowerPoint, and cross-application workflows through canonical WPS Actions, a bundled Python 3.9+ standard-library Runner, and a bundled WPS Add-in.

## Use

Load [`skills/wps-office/SKILL.md`](skills/wps-office/SKILL.md). It guides the agent to:

1. run `python3 scripts/wps.py check` from the skill directory;
2. read the Action manifest and the relevant progressive reference;
3. invoke exactly one WPS Action with `python3 scripts/wps.py invoke '<request-json>'`.

The Runner installs or updates the Add-in in the current user profile on Linux, macOS, and Windows. Restart WPS Office when `check` returns `restart_required`.

## Safety

The manifest classifies every Action as `read`, `write`, or `destructive`. The Runner rejects destructive Actions unless the request contains `"confirmed": true` after a specific user confirmation. It binds only to loopback and uses a per-user credential for Add-in communication.

## Verification

```bash
python3 scripts/validate_action_manifest.py
python3 -m unittest discover -s tests
```

The checked-in migration ledger records the retired historical capabilities and their canonical replacements. Platform certification remains separate from migration completion: a platform is only certified after its real-WPS smoke tests pass.
