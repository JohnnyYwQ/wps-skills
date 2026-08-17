---
status: accepted
---

# Install the WPS add-in on first use

The WPS Skill Package will include an idempotent Python installer that detects, installs, and updates the unified WPS Add-in in the current user's WPS profile before the first action. It will merge its registration into `publish.xml` without overwriting other add-ins and ask the user to restart WPS only when installation or an update requires it.

## Consequences

Loading the skill alone does not imply that WPS has loaded the add-in. Every action workflow must run a readiness check first, and the installer must handle the user-level Windows and Linux WPS add-in locations without requiring administrator privileges.
