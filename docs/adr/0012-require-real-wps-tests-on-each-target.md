---
status: accepted
---

# Require real WPS tests on every supported target

Support claims require behavioral smoke tests against an installed WPS on Linux x86_64, Linux ARM64, Windows x86_64, and Windows ARM64. Static checks, mocked protocol tests, or success on one architecture are necessary but insufficient substitutes for any matrix entry.

## Consequences

A supported release requires access to all four environments before release. Each environment must exercise add-in installation and readiness plus representative Excel, Word, PowerPoint, common, destructive-confirmation, save, and cross-application workflows. A Migration Complete preview, as defined by ADR-0015, is not a supported release and does not grant Platform Certified status to any target.
