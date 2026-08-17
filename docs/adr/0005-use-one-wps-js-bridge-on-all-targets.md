---
status: accepted
---

# Use one WPS JavaScript bridge on all target platforms

Linux x86_64, Linux ARM64, Windows x86_64, and Windows ARM64 will all deliver WPS Actions through the WPS-hosted JavaScript add-in and a loopback polling protocol. The Windows PowerShell COM path is retired because it is architecture-sensitive, unverified on Windows ARM64, and would preserve a second platform implementation.

## Consequences

The JavaScript add-in currently dispatches 227 actions while the PowerShell implementation dispatches 241. The 14 missing actions must be implemented in the add-in and included in behavioral parity tests before the unified bridge can claim full capability equivalence.
