---
status: accepted
---

# Use Python 3.9 or newer with no third-party packages

Bundled Python scripts will support Python 3.9 or newer and use only the standard library. The skill must not require `pip install`, architecture-specific wheels, or Python COM and HTTP packages; the same Python source is distributed on Linux x86_64, Linux ARM64, Windows x86_64, and Windows ARM64.

## Consequences

Loopback HTTP, JSON transport, XML registration, file copying, locking, timeouts, and process detection must be implemented with the standard library. Platform acceptance still requires a compatible Python interpreter and WPS installation on each target, but the skill itself has no per-architecture Python build artifact.
