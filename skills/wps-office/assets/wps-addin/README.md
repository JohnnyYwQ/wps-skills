# Unified WPS Add-in

此目录是 WPS Skill Package 随包分发的统一 Platform Bridge 资源，供 Linux 与 Windows 的 x86_64、ARM64 目标使用。

- `manifest.xml`、`ribbon.xml`：WPS Add-in 注册和功能区入口。
- `main.js`：轮询 Python Runner 并分发 WPS Action。
- `handlers/`、`utils/`：现有 WPS Action 行为实现与响应辅助代码。
- `wps-skill-config.js`：安装时生成，不进入源目录；保存当前用户的共享本机认证凭证。

修改资源、加载顺序或认证请求头时，应同步更新此文件和 `skills/wps-office/README.md`。
