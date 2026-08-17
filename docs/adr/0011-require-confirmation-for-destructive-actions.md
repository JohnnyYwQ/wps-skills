---
status: accepted
---

# Require confirmation for destructive WPS Actions

Every action manifest entry is classified as `read`, `write`, or `destructive`. Read actions require no confirmation, ordinary writes may proceed when the user's request already authorizes the change, and destructive actions such as deletion, overwrite, or discarding unsaved work require explicit user confirmation before invocation.

## Consequences

The skill owns the conversational confirmation workflow and passes an explicit confirmation marker to the Python runner. The runner rejects destructive actions without that marker, so safety does not depend only on the agent following prose instructions.
