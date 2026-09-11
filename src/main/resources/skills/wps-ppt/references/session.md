# 通过包内命令持有一个 PPT Action Session

在安装了 WPS 的 Windows 主机上调用本 Skill 的 `scripts/ppt.py`。需要可见窗口时，在已登录的用户桌面执行；SSH 不保证窗口出现在该桌面。仅在任务要求远程执行时使用已有授权的桌面启动机制。Windows 路径不会自动映射为 Mac 路径。

Agent 直接运行下列入口，不编写 Python、PowerShell 任务脚本或内联代码来管理 Session。复杂参数写成 UTF-8 JSON 数据文件即可；文件只包含一个 Action 的参数，不包含动作列表。每个命令可以由一次独立终端工具调用运行，无需维持 REPL 或 stdin。

## 开始与逐次执行

先按 SKILL.md 解析本次需要的完整契约，再启动一次：

```powershell
python "<skill-dir>/scripts/ppt.py" --app ppt --start
```

返回 `handle`、`nextStep: 1`、`ready`（含底层 Session 身份和日志路径）。记录原样返回的 `handle`，只在当前任务中复用。它定位包内持有的 Session Client，不是文档地址。`--start` 不创建或打开文档。

仅当用户要求新建时，第一步调用：

```powershell
python "<skill-dir>/scripts/ppt.py" --app ppt --call <handle> --step 1 --action createPresentation
```

已有文件任务的第一步改为本应用的 open Action，并传用户明确指定的绝对文件路径。省略参数选项表示 `{}`，不表示缺失的契约参数会自动补齐。

阅读返回的 `response`，确认 `outcome`、`data` 或 `error` 以及 `canExecute`，再选择下一步；`--step` 使用上一步返回的 `nextStep`。例如 `listSlides` 的参数文件可为：

```json
{}
```

将数据保存为 `params.json`，然后单独调用：

```powershell
python "<skill-dir>/scripts/ppt.py" --app ppt --call <handle> --step 2 --action listSlides --params-file "<absolute-path>/params.json"
```

示例只演示入口，不代表完整任务或最终验证。后续写入、检查、保存分别使用解析过的 Action 和新的步骤编号。观察 token、内容范围、工作表/幻灯片标识必须来自这个会话的实际响应。

参数也可通过 `--params-stdin` 传一个 JSON 对象。`--params-json` 适用于能可靠保留 JSON 引号的终端；Windows PowerShell 优先使用 `--params-file`，避免本地代码页和原生命令引号转换。参数文件支持 UTF-8 BOM。不要把多个命令连成无人检查响应的批处理。

## 返回结果与恢复

- `response` 是完整 Action Response；数据在 `response.data`，实际失败在 `response.error`，不会被清理错误覆盖。
- `canExecute` 和 `nextStep` 描述回执生成时的状态。历史回执不证明会话现在仍可用；当前状态用 `--status`。
- `sessionOutcome` 与 `cleanupError` 单独报告资源清理。顶层 `error` 描述命令、连接或生命周期问题，不冒充 Action Response。
- 正常命令退出码为 0；Action 的 `failed`/`unknown` 为 2；命令、通道或清理问题为 4。失败后仍应读取 JSON，不能只看退出码。`--status`/`--close` 不因历史 Action 失败而改写清理结果。

查询状态或某一步的既有回执：

```powershell
python "<skill-dir>/scripts/ppt.py" --app ppt --status <handle>
python "<skill-dir>/scripts/ppt.py" --app ppt --status <handle> --step 2
```

同一步的相同 Action 和参数只取回已有结果或等待正在执行的结果，不再次执行；同一步改换请求返回 `STEP_CONFLICT`。一次只能有一个 Action 在执行，不能提前提交下一步。

`COMMAND_WAIT_TIMEOUT` 表示命令等待结束，不表示 WPS 停止执行。先查状态或用完全相同的步骤和参数取回回执，不能换步骤号重新写入。`mayHaveEffect: true` 且没有当前响应时，效果尚不能确定，不能拿上一条成功响应替代当前结果。可用会话中的 `failed` 或 `unknown` 由 Agent 评估后显式选择检查或结束；终止的会话不自动重启、重绑或重建文档。

## 结束、超时与本地回执

任务完成或取消后直接结束：

```powershell
python "<skill-dir>/scripts/ppt.py" --app ppt --close <handle>
```

关闭可重复调用；若已有 Action 正在执行，先完成其结果处理再清理。结束不保存或关闭 WPS 文档。新建未保存文档也会保留打开；需保存时必须先显式执行 `saveAs` 并验证。

包内 Session Client 在两个命令之间持有原有 Host 通道。300 秒没有新 Action 时回收；查询状态和重复取回回执不会续期。调用终端退出不结束这个包内 Client；Client 本身异常退出会使 Host 失去通道并清理，原有 Lease/Quarantine 规则继续适用。孤立 Client 最迟由空闲期限回收，失联 handle 不会创建替代 Session。

`--start --timeout 90` 在启动前设置 Host 启动和每个 Action 的等待上限；后续 `--call --timeout 90` 只设置本次命令等待回执的上限。默认均为 60 秒；超时不授权重放。

Windows 本地回执目录为 `%LOCALAPPDATA%\WpsSkills\sessions`，可在启动前用 `WPS_SKILLS_SESSION_DIR` 指定本机用户私有目录，并让后续命令使用相同配置。目录需要支持原子替换和硬链接，不使用共享网络目录。每个 handle 目录保留参数、结果和 `worker.stderr.log`，用于命令丢失后的查询；任务结束并完成诊断后可以删除对应目录，不删除仍在运行的会话目录。标准 Action/Session 日志见 [logging.md](logging.md)。

原有 Python Client 和原始 `--session` JSONL 入口仅供现有集成调用。Agent 文档任务使用上述包内管理入口。
