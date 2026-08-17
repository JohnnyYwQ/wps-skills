---
status: accepted
---

# Use WPS Action identifiers as the capability contract

Names such as `getCellValue` are the canonical capability identifiers for the final skill. Legacy MCP tool names are retained only in migration analysis and are not part of the distributed skill interface; one machine-readable action manifest is the source of truth for parameters, results, application ownership, prerequisites, mutation risk, documentation, and coverage tests.

## Consequences

The Python runner must reject actions absent from the manifest and perform its standard-library validation before contacting WPS. Reference documentation and capability counts derive from the manifest, and tests must verify that every declared action is dispatched by the unified JavaScript add-in.
