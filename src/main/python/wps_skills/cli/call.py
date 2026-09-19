"""Repository entry point for application Task submission."""
from wps_skills.cli.task import main as _main


def main(argv=None, **kwargs):
    return _main(argv, required_application=None, **kwargs)


if __name__ == "__main__":
    raise SystemExit(main())
