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
