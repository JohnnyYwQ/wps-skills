# WPS Skill Package

此目录是智能体可加载的 WPS Skill Package。

- `SKILL.md`：选择和调用只读 WPS Action 的核心流程。
- `references/action-manifest.json`：WPS Action 名称、参数、结果、前置条件与风险的唯一事实来源。
- `scripts/wps.py`：Python 3.9+ 标准库 Runner；每个进程通过临时 loopback HTTP 服务执行一个 WPS Action，然后释放端口并退出。

修改目录结构、Runner 进程契约或 WPS Action 调用流程时同步更新此文件。
