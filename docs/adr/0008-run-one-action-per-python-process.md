---
status: accepted
---

# Run one WPS Action per Python process

Each agent invocation will start the bundled Python runner, perform add-in readiness checks, bind the loopback polling endpoint, deliver one WPS Action, return one JSON result, and exit. Calls are serialized with a cross-platform file lock so no user-managed daemon or port-sharing race is introduced.

## Consequences

Multi-step workflows are orchestrated by the agent through sequential script invocations. Standard output is reserved for the machine-readable JSON result and diagnostics go to standard error; state that must survive between actions cannot rely on process memory.
