# WPS Skill Package

此目录是智能体可加载的 WPS Skill Package。

- `SKILL.md`：安装、就绪检查、选择和调用 WPS Action，以及写入和破坏性操作风险门禁的核心流程。
- `references/action-manifest.json`：WPS Action 名称、参数、结果、前置条件与风险的唯一事实来源。
- `references/excel.md`：Excel 核心数据以及格式、图表、分析、透视、保护、图片和上下文工作流的渐进式 Action 指南。
- `references/word.md`：Word 文档生命周期、内容、格式、模板、书签、批注和修订工作流的渐进式 Action 指南。
- `references/powerpoint.md`：PowerPoint 演示文稿、幻灯片、文本、图片、表格、备注、基础版式和高级设计工作流的渐进式 Action 指南。
- `scripts/wps.py`：Python 3.9+ 标准库 Runner；提供 `check`、`install` 和 `invoke` 内部操作。
- `scripts/wps_skill/addon_installer.py`：在 Linux、macOS、Windows 用户配置目录幂等安装 WPS Add-in、安全合并 `publish.xml`、清理 macOS 目标目录的 Gatekeeper 隔离属性，并管理本机认证配置。
- `assets/wps-addin/`：Linux x86_64、Linux ARM64、Windows x86_64 和 Windows ARM64 共用的 WPS Add-in 资源。

WPS Add-in 的浏览器入口和 manifest 都必须先加载安装器生成的 `wps-skill-config.js`，再加载 `main.js`，使 loopback 轮询使用与 Runner 相同的本机凭据和安装摘要。

`check` 会在首次使用或资源变化时原子安装 Add-in，并返回 `ready`、`restart_required`、`wps_not_running` 或 `addin_unavailable`。macOS 安装到 WPS 容器内的 `jsaddons` 目录，并在复制后清理目标 Add-in 的 Gatekeeper 隔离属性；仍须完全退出并重新打开 WPS 才能加载更新。非就绪结果可在 `data.error` 中附带稳定的传输错误，同时保留兼容的 readiness 状态。安装摘要会保持 `restart_required`，直到 WPS 中已加载的 Add-in 通过认证 ping 回报相同摘要。认证凭证只写入用户配置与已安装 Add-in，不通过标准输出返回。

`check` 与 `invoke` 使用当前用户配置目录内的跨进程系统文件锁，竞争调用立即返回可重试的 `ACTION_BUSY`。Runner 只绑定 `127.0.0.1`，所有 Action 轮询与结果请求均使用本机共享凭证认证，并在成功、协议错误、断开、超时或端口冲突后关闭临时服务并释放锁。浏览器 CORS `OPTIONS` 预检无法携带认证头，因此只返回空协商响应，不读取或改变 Action 状态。锁由操作系统持有，进程退出后遗留的空锁文件不会继续占锁。

Runner 从 Action 清单读取 `read`、`write` 或 `destructive` 风险等级。只读和普通写入 Action 可按正常流程执行；破坏性 Action 只有在请求包含布尔值 `"confirmed": true` 时才会送达 WPS Add-in，否则在传输前返回 `CONFIRMATION_REQUIRED`。缺失或未知风险等级返回 `INVALID_ACTION_RISK`，同样不会接触 WPS。

Action 传输错误使用稳定分类：未轮询为 `ADDIN_NOT_READY`，已投递后超时为 `ACTION_TIMEOUT`，结果上传断开为 `ADDIN_DISCONNECTED`，无效 JSON 为 `INVALID_ADDIN_JSON`，非对象结果信封为 `INVALID_ADDIN_RESPONSE`，请求标识不匹配为 `REQUEST_ID_MISMATCH`，端口占用为 `PORT_IN_USE`，其他端口打开失败为 `PORT_UNAVAILABLE`。诊断不得记录共享凭证或 Action 参数。

修改目录结构、Runner 进程契约或 WPS Action 调用流程时同步更新此文件。
