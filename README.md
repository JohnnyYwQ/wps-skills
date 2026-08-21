# WPS Office Skill

`skills/wps-office/` 是本仓库唯一面向智能体的 WPS Skill Package。智能体通过规范的 WPS Actions、随包提供的 Python 3.9+ 标准库 Runner 和 WPS Add-in，操作 Excel、Word、PowerPoint 及跨应用工作流。

## 使用

加载 [`skills/wps-office/SKILL.md`](skills/wps-office/SKILL.md)，并按照其中的流程：

1. 在 Skill 目录运行 `python3 scripts/wps.py check`；
2. 阅读 Action manifest 和相关的渐进式参考文档；
3. 使用 `python3 scripts/wps.py invoke '<request-json>'` 调用一个 WPS Action。

Runner 会在 Linux、macOS 和 Windows 的当前用户配置中安装或更新 Add-in。若 `check` 返回 `restart_required`，请完全退出并重新启动 WPS Office。

## 安全

manifest 将每个 Action 标记为 `read`、`write` 或 `destructive`。对于破坏性 Action，必须先获得用户针对具体后果的明确确认，并在请求中包含 `"confirmed": true`，否则 Runner 会拒绝执行。Runner 只绑定本机回环地址，并使用当前用户的本机凭证与 Add-in 通信。

## 验证

```bash
python3 scripts/validate_action_manifest.py
python3 -m unittest discover -s tests
```

仓库中的迁移台账记录了已退役的历史能力及其规范替代方案。平台认证与迁移完成相互独立：只有在对应平台的真实 WPS smoke tests 通过后，才能认定该平台已通过认证。

## 开发

实现架构、事实来源地图，以及增加、修改或删除仓库功能的操作指南，见 [`DEVELOPMENT.md`](DEVELOPMENT.md)。
