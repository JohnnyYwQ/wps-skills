---
status: accepted
---

# Distribute the integration as an agent skill package

The deliverable is a self-contained WPS Skill Package loaded by one AI agent, not a separately operated CLI application or daemon product. The target agent is guaranteed to be able to execute local Python scripts. `SKILL.md` will direct the agent's workflow, bundled Python scripts will perform deterministic WPS communication, and detailed WPS Action contracts will be loaded from references only when needed.

## Consequences

The package may expose an internal command-line contract for the agent to execute, but it is not a user-facing product interface. The TypeScript catalog of hundreds of MCP-shaped tool definitions does not need a line-for-line Python port when the same capability can be expressed through concise skill instructions, action references, and a generic Python runner.
