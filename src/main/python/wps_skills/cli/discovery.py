"""Side-effect-free discovery of the selected Application Contract Set."""

import importlib
import importlib.util
import json


def discover(args, output_stream, error_stream):
    from wps_skills.core.action_runtime import ActionAddress
    if importlib.util.find_spec("wps_skills." + args.app) is None:
        error_stream.write(f"WPS_DISCOVERY_UNAVAILABLE app={args.app}\n")
        return 4
    module = importlib.import_module("wps_skills." + args.app + ".contracts")
    contracts = getattr(module, args.app.upper() + "_PRODUCTION_CONTRACT_SET", None)
    if contracts is None:
        error_stream.write(f"WPS_DISCOVERY_UNAVAILABLE app={args.app}\n")
        return 4
    if args.index:
        value = {
            "app": args.app,
            "actions": [entry.to_wire() for entry in contracts.action_index()],
        }
    else:
        value = contracts.batch_resolve([
            ActionAddress(app=args.app, action=action) for action in args.resolve
        ])
    output_stream.write(json.dumps(value, ensure_ascii=True, indent=2) + "\n")
    return 0 if args.index or value["status"] == "complete" else 2


