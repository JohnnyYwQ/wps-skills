"""Assemble the standalone Ppt Application Skill."""

import argparse
from pathlib import Path
from wps_skills.cli.build_skill import build_application_skill


def build_ppt_skill(destination):
    return build_application_skill("ppt", destination)


def main(argv=None):
    parser = argparse.ArgumentParser(description="Assemble the complete wps-ppt Skill directory")
    parser.add_argument("--output", type=Path, default=Path("build/skills/wps-ppt"))
    args = parser.parse_args(argv)
    try:
        print(build_ppt_skill(args.output))
    except FileExistsError as exc:
        parser.exit(2, str(exc) + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
