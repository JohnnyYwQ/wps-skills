"""Assemble the standalone Excel Application Skill."""

import argparse
from pathlib import Path
from wps_skills.cli.build_skill import build_application_skill


def build_excel_skill(destination):
    return build_application_skill("excel", destination)


def main(argv=None):
    parser = argparse.ArgumentParser(description="Assemble the complete wps-excel Skill directory")
    parser.add_argument("--output", type=Path, default=Path("build/skills/wps-excel"))
    args = parser.parse_args(argv)
    try:
        print(build_excel_skill(args.output))
    except FileExistsError as exc:
        parser.exit(2, str(exc) + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
