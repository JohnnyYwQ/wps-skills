---
status: accepted
---

# Separate migration completion from platform certification

The repository may become Migration Complete and merge or distribute a clearly labeled preview before any target is Platform Certified. Keeping the Legacy MCP Runtime until all four real-machine environments are available would couple source migration to external hardware access without increasing confidence in the final runtime. ADR-0012's real-WPS test gate remains in force for a supported release: a Migration Complete preview is not a supported release, and no platform may be described as supported until its own WPS smoke tests pass.

## Migration Complete gate

- The WPS Skill Package is self-contained and has no Node.js, TypeScript, MCP, or legacy-source dependency.
- The explicit retirement ledger is applied, contains no unresolved conflicts, and preserves corrected mappings.
- The Action manifest and packaged JavaScript Add-in dispatch set are identical.
- Validators read final artifacts and frozen migration evidence rather than executable legacy sources.
- Contract, installer, and Fake Add-in tests pass reliably.
- The Legacy MCP Runtime, old MCP configuration and documentation, and independently loadable legacy WPS skills are removed.

## Consequences

Migration Complete builds may enter the main branch or be distributed as previews, but documentation must distinguish them from supported releases. Platform Certified is awarded separately to each operating-system and architecture pair only after the ADR-0012 smoke tests pass; certification work does not restore the Legacy MCP Runtime.
