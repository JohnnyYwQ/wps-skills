"""Assemble the standalone Word Application Skill."""

import argparse
from pathlib import Path
from wps_skills.cli.build_skill import MAIN, build_application_skill, write_word_action_definitions


def build_word_skill(destination):
    return build_application_skill("word", destination)


def main(argv=None):
    parser = argparse.ArgumentParser(description="Assemble the complete wps-word Skill directory")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--output", type=Path, default=Path("build/skills/wps-word"))
    mode.add_argument("--refresh-actions", action="store_true", help="Regenerate Word contract references and per-Action query schemas without building a package")
    args = parser.parse_args(argv)
    if args.refresh_actions:
        print(write_word_action_definitions(MAIN / "resources" / "skills" / "wps-word"))
        return 0
    try:
        print(build_word_skill(args.output))
    except FileExistsError as exc:
        parser.exit(2, str(exc) + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
