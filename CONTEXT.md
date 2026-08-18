# WPS Automation

This context describes the language used when exposing WPS Office capabilities to AI callers across supported operating systems.

## Language

**WPS Tool**:
An AI-callable WPS capability with a stable name, input contract, and result contract.
_Avoid_: MCP tool, command

**WPS Action**:
A platform-neutral instruction to perform one concrete operation in WPS Office.
_Avoid_: Tool, native call

**Platform Bridge**:
The platform-specific path that delivers a WPS Action to WPS Office and returns its result.
_Avoid_: MCP bridge, tool engine

**WPS Skill Package**:
A self-contained folder that an AI agent loads to learn WPS workflows and gain access to the bundled execution resources.
_Avoid_: Standalone application, MCP server

**WPS Add-in**:
WPS-hosted JavaScript that receives WPS Actions from the skill's Python runner and executes them inside WPS Office.
_Avoid_: Skill, MCP server

**Legacy MCP Runtime**:
The retired Node.js, TypeScript, MCP, and PowerShell execution path retained only while its stable WPS behavior is extracted into the WPS Skill Package.
_Avoid_: WPS Skill Package, fallback runtime

**Retired Legacy Capability**:
A former MCP-facing behavior deliberately excluded from the WPS Skill Package because it was runtime-only, had a broken or ambiguous contract, or is covered by canonical WPS Actions.
_Avoid_: Missing Action, unsupported platform

**Migration Complete**:
The repository state in which the WPS Skill Package is self-contained, its automated contracts pass, and the Legacy MCP Runtime has been removed. It makes no claim that WPS behavior has been verified on a particular operating-system and architecture pair.
_Avoid_: Platform support, release certified

**Platform Certified**:
A specific operating-system and architecture pair whose installed WPS has passed the required behavioral smoke tests for the WPS Skill Package.
_Avoid_: Migration complete, statically compatible
