---
status: accepted
---

# Recognize macOS Add-in installation as experimental

The WPS Skill Package will recognize macOS x86_64 and ARM64 for Add-in installation in WPS's container profile and clear the copied Add-in's Gatekeeper quarantine attribute. macOS remains outside the supported-target matrix until behavioral smoke tests against a matching WPS distribution demonstrate installation, readiness, and representative WPS Actions; this preserves the four accepted Linux/Windows targets while making the macOS failure mode diagnosable.
