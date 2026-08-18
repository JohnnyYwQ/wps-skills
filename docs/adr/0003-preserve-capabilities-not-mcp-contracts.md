---
status: superseded by ADR-0014
---

# Preserve WPS capabilities rather than MCP contracts

The WPS Skill Package must provide behaviorally equivalent Excel, Word, PowerPoint, common, and cross-application capabilities. MCP discovery, MCP tool definitions, MCP response wrappers, and MCP-specific names are not compatibility requirements; the stable compatibility seam is the existing WPS Action contract used by the platform bridges.

## Consequences

Migration completeness will be measured by a generated capability-to-action manifest and behavioral tests, not by a line-for-line Python translation of the TypeScript handlers or identical human-readable response text.
