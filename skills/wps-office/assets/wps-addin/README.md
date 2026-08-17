# Unified WPS Add-in

此目录是 WPS Skill Package 随包分发的统一 Platform Bridge 资源，供 Linux 与 Windows 的 x86_64、ARM64 目标使用。

- `manifest.xml`、`ribbon.xml`：WPS Add-in 注册和功能区入口。
- `main.js`：轮询 Python Runner、分发 WPS Action 并包含现有行为实现。Excel 切片覆盖核心数据以及格式、图表、分析、透视、保护、图片和上下文能力；Word 切片覆盖文档生命周期、内容、范围、模板、格式、书签、批注与修订能力。成功结果遵循 Action manifest，WPS 异常返回失败结果。
- `wps-skill-config.js`：安装时生成，不进入源目录；保存当前用户的共享本机认证凭证与安装摘要。

上述 Word 覆盖表示统一 Add-in 已包含对应实现和模拟协议测试，不代表四个目标平台已经完成支持验收。平台支持声明仍需 ADR-0012 要求的对应真实 WPS smoke test；这些验收由后续平台 Issue 交付。

修改资源、加载顺序或认证请求头时，应同步更新此文件和 `skills/wps-office/README.md`。旧运行时中的拆分 handler 未被清单加载，因此不进入该分发目录。
