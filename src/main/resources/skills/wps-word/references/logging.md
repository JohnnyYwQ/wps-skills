# 运行日志

日志写在运行 Session Host 的机器上，不在 Skill 安装目录，也不在文档输出目录。Windows 默认位置：

```text
%LOCALAPPDATA%\wps-skills\logs\
├── sessions\session-<id>.jsonl
└── actions\trace-<id>.jsonl
```

三端共用此根目录，每个会话／动作有独立文件；每条记录包含 `app` 和 `sessionId`，动作记录另有 `traceId`。日志记录执行事件、结果和耗时。

在启动 Session 前设置自定义根目录：

```powershell
$env:WPS_TRACE_DIR = "D:\wps-logs"
```

也可以由启动 Python 的环境传入 `WPS_TRACE_DIR`。默认位置不可用时会依次尝试其他位置，最终回退到 `%TEMP%\wps-skills-logs`；全部不可写时允许无日志继续。以 `client.ready["traceLog"]` 和每次响应的 `traceLog` 为准，`null` 表示没有分配可用日志路径。能力发现 `--index`／`--resolve` 不创建运行日志。

日志目前没有自动轮转或保留期限。另一个目录 `%LOCALAPPDATA%\WpsSkills\document-coordination` 是 Document Lease／Quarantine 协调状态，不是日志，不作为日志清理对象。
