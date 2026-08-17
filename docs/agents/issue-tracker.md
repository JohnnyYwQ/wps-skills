# Issue tracker: GitHub

Issues and specs for this repo live as GitHub issues. Use the `gh` CLI for all operations.

## Repository

Infer the repository from the Git remote when working inside the clone. The configured project repository is `JohnnyYwQ/wps-skills`.

## Conventions

- Create an issue with `gh issue create`.
- Read an issue and its discussion with `gh issue view <number> --comments`.
- List issues with `gh issue list` and request JSON output when filtering is needed.
- Comment with `gh issue comment <number>`.
- Apply or remove labels with `gh issue edit <number> --add-label <label>` or `--remove-label <label>`.
- Close an issue with `gh issue close <number> --comment <message>`.

## Pull requests as a triage surface

**PRs as a request surface: no.**

## Publishing

When an engineering skill says to publish a spec or ticket to the issue tracker, create a GitHub issue in this repository.
