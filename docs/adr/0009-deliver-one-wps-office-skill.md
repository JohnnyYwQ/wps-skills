---
status: accepted
---

# Deliver one WPS Office skill

The final package contains one `wps-office` skill and one triggering `SKILL.md`; the existing Excel, Word, PowerPoint, and cross-application skills will not remain independently loadable. The root instructions own routing, execution, readiness, and safety, while application-specific action guidance is split into directly linked `references/` files and loaded only when relevant.

## Consequences

Nested skill discovery is not part of the design. Duplicated instructions and MCP tool lists must be consolidated, and the final `SKILL.md` must remain concise enough to load for every WPS task.

Plugin or marketplace metadata may remain as an optional installation mechanism, but it must expose only the single `wps-office` skill and must not introduce another runtime interface. The bundled Python Runner and WPS Add-in are internal resources of that skill, not independently distributed products.
