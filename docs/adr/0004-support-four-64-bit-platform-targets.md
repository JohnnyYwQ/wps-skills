---
status: accepted
---

# Support four 64-bit platform targets

The WPS Skill Package will be accepted on Linux x86_64, Linux ARM64, Windows x86_64, and Windows ARM64. Windows ia32 is explicitly outside the support matrix, so Python execution, packaging, WPS installation, and smoke tests may assume a 64-bit process.

## Consequences

Each operating-system and CPU-architecture pair is a separate acceptance target. Passing on x86_64 does not imply ARM64 support, and platform claims require a matching WPS distribution plus behavioral smoke tests on that target.
