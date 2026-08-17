---
status: superseded by ADR-0005
---

# Use a Python runtime while retaining the existing WPS-side bridges

The production installation must not require Node.js. The WPS Skill Package will own the capability catalog, while bundled Python will own action invocation, platform selection, local polling, and process state; the existing Windows PowerShell COM implementation and WPS-hosted JavaScript add-in remain because they execute inside platform-specific environments and do not require Node.js.

## Consequences

Keeping the TypeScript tool engine behind a Python wrapper is not an accepted final architecture. Existing WPS Action names, parameters, and results must remain stable so the platform-specific implementations can be reused without being rewritten.
