# Unified WPS Add-in

此目录是 WPS Skill Package 随包分发的统一 Platform Bridge 资源，供 Linux 与 Windows 的 x86_64、ARM64 目标使用。

- `manifest.xml`、`ribbon.xml`：WPS Add-in 注册和功能区入口。
- `main.js`：轮询 Python Runner、分发 WPS Action 并包含现有行为实现。Excel 核心切片覆盖工作簿、工作表、单元格、区域、公式和基础数据处理，成功结果遵循 Action manifest，WPS 异常返回失败结果。
- `wps-skill-config.js`：安装时生成，不进入源目录；保存当前用户的共享本机认证凭证与安装摘要。

修改资源、加载顺序或认证请求头时，应同步更新此文件和 `skills/wps-office/README.md`。旧运行时中的拆分 handler 未被清单加载，因此不进入该分发目录。
