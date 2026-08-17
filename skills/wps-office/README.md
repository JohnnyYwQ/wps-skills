# WPS Skill Package

此目录是智能体可加载的 WPS Skill Package。

- `SKILL.md`：安装、就绪检查、选择和调用只读 WPS Action 的核心流程。
- `references/action-manifest.json`：WPS Action 名称、参数、结果、前置条件与风险的唯一事实来源。
- `scripts/wps.py`：Python 3.9+ 标准库 Runner；提供 `check`、`install` 和 `invoke` 内部操作。
- `scripts/wps_skill/addon_installer.py`：在 Linux/Windows 用户配置目录幂等安装 WPS Add-in、安全合并 `publish.xml` 并管理本机认证配置。
- `assets/wps-addin/`：Linux x86_64、Linux ARM64、Windows x86_64 和 Windows ARM64 共用的 WPS Add-in 资源。

`check` 会在首次使用或资源变化时原子安装 Add-in，并返回 `ready`、`restart_required`、`wps_not_running` 或 `addin_unavailable`。认证凭证只写入用户配置与已安装 Add-in，不通过标准输出返回。

修改目录结构、Runner 进程契约或 WPS Action 调用流程时同步更新此文件。
