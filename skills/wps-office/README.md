# WPS Skill Package

此目录是智能体可加载的 WPS Skill Package。

- `SKILL.md`：安装、就绪检查、选择和调用只读 WPS Action 的核心流程。
- `references/action-manifest.json`：WPS Action 名称、参数、结果、前置条件与风险的唯一事实来源。
- `scripts/wps.py`：Python 3.9+ 标准库 Runner；提供 `check`、`install` 和 `invoke` 内部操作。
- `scripts/wps_skill/addon_installer.py`：在 Linux/Windows 用户配置目录幂等安装 WPS Add-in、安全合并 `publish.xml` 并管理本机认证配置。
- `assets/wps-addin/`：Linux x86_64、Linux ARM64、Windows x86_64 和 Windows ARM64 共用的 WPS Add-in 资源。

`check` 会在首次使用或资源变化时原子安装 Add-in，并返回 `ready`、`restart_required`、`wps_not_running` 或 `addin_unavailable`。非就绪结果可在 `data.error` 中附带稳定的传输错误，同时保留兼容的 readiness 状态。安装摘要会保持 `restart_required`，直到 WPS 中已加载的 Add-in 通过认证 ping 回报相同摘要。认证凭证只写入用户配置与已安装 Add-in，不通过标准输出返回。

`check` 与 `invoke` 使用当前用户配置目录内的跨进程系统文件锁，竞争调用立即返回可重试的 `ACTION_BUSY`。Runner 只绑定 `127.0.0.1`，所有 Action 轮询与结果请求均使用本机共享凭证认证，并在成功、协议错误、断开、超时或端口冲突后关闭临时服务并释放锁。浏览器 CORS `OPTIONS` 预检无法携带认证头，因此只返回空协商响应，不读取或改变 Action 状态。锁由操作系统持有，进程退出后遗留的空锁文件不会继续占锁。

Action 传输错误使用稳定分类：未轮询为 `ADDIN_NOT_READY`，已投递后超时为 `ACTION_TIMEOUT`，结果上传断开为 `ADDIN_DISCONNECTED`，无效 JSON 为 `INVALID_ADDIN_JSON`，非对象结果信封为 `INVALID_ADDIN_RESPONSE`，请求标识不匹配为 `REQUEST_ID_MISMATCH`，端口占用为 `PORT_IN_USE`，其他端口打开失败为 `PORT_UNAVAILABLE`。诊断不得记录共享凭证或 Action 参数。

修改目录结构、Runner 进程契约或 WPS Action 调用流程时同步更新此文件。
