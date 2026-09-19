"""Assemble the standalone PPT Application Skill."""

import argparse
from pathlib import Path
from wps_skills.cli.build_skill import MAIN, build_application_skill, write_action_definitions


def build_ppt_skill(destination):
    return build_application_skill("ppt", destination)


def main(argv=None):
    parser = argparse.ArgumentParser(description="Assemble the complete wps-ppt Skill directory")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--output", type=Path, default=Path("build/skills/wps-ppt"))
    mode.add_argument("--refresh-actions", action="store_true", help="Regenerate PPT contract references and per-Action query schemas without building a package")
    args = parser.parse_args(argv)
    if args.refresh_actions:
        print(write_action_definitions("ppt", MAIN / "resources" / "skills" / "wps-ppt"))
        return 0
    try:
        print(build_ppt_skill(args.output))
    except FileExistsError as exc:
        parser.exit(2, str(exc) + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
