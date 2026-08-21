# WPS Automation

This context describes the language used when exposing WPS Office capabilities to AI callers across supported operating systems.

## Language

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

**Agent Host**:
The AI application that discovers and loads a WPS Skill Package and executes its bundled runner. During a real-WPS test, the Agent Host runs on the Validation Target so the runner and WPS Add-in share the target's loopback interface and user profile.
_Avoid_: Control host, remote bridge

**Readiness Check**:
An Agent Host-initiated check before a WPS Action workflow that ensures the current WPS environment can accept WPS Actions, potentially requiring WPS Office to be started or restarted.
_Avoid_: Environment Preflight, installation check

**Agent Host Smoke Test**:
A candidate-specific test in which an Agent Host discovers and loads an unchanged WPS Skill Package, resolves its bundled resources, and drives an observable WPS result from a realistic user request. It demonstrates host integration but does not grant Platform Certified status.
_Avoid_: Runner test, environment preflight, platform certification

**Control Host**:
The operator machine that stages candidates, orchestrates remote checks, and collects sanitized records from Validation Targets. It does not proxy WPS Action loopback traffic for another machine.
_Avoid_: Agent Host, platform bridge

**Validation Target**:
The machine or virtual machine whose operating system and architecture are under test and where the Agent Host, Python runner, WPS Add-in, and WPS Office execute together.
_Avoid_: Control host, inferred platform

**Legacy MCP Runtime**:
The retired Node.js, TypeScript, MCP, and PowerShell execution path retained only while its stable WPS behavior is extracted into the WPS Skill Package.
_Avoid_: WPS Skill Package, fallback runtime

**Retired Legacy Capability**:
A former MCP-facing behavior deliberately excluded from the WPS Skill Package because it was runtime-only, had a broken or ambiguous contract, or is covered by canonical WPS Actions.
_Avoid_: Missing Action, unsupported platform

**WPS Action Retirement**:
The deliberate removal of a published WPS Action through a GitHub Issue and ADR that record its impact, replacement or migration path, and coordinated contract changes.
_Avoid_: Direct deletion, cleanup

**Migration Complete**:
The repository state in which the WPS Skill Package is self-contained, its automated contracts pass, and the Legacy MCP Runtime has been removed. It makes no claim that WPS behavior has been verified on a particular operating-system and architecture pair.
_Avoid_: Platform support, release certified

**Environment Preflight**:
The platform-specific check that a real WPS validation environment is accessible, correctly provisioned, resettable, and ready to receive a WPS Skill Package candidate. It does not certify WPS Skill Package behavior on that platform.
_Avoid_: Platform validation, platform acceptance, certification

**Platform Certified**:
A specific operating-system and architecture pair whose installed WPS has passed the required behavioral smoke tests for the WPS Skill Package.
_Avoid_: Migration complete, statically compatible
