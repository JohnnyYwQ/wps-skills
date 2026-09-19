# 启动与通信实验：里程碑 1

从 Mac 终端打包当前工作区，在 Windows 已登录桌面独立运行测试，随后查询或取回证据。当前只测 **现有启动链 + stdin/stdout 匿名管道**，不包含启动简化、命名管道、TCP 的比较结果，也不替换已安装的 Skill。

## 开始使用

要求本机 Python 3.10+、SSH/scp；SSH alias `win` 指向测试 Windows 用户。Windows 需要已登录的交互桌面、Python 3.10+、Windows PowerShell 和 WPS 表格。不使用 SSH Session 0 执行 COM。

在仓库根目录：

```sh
python3 scripts/debug/startup_communication.py pack
```

命令打印新的本地运行目录。将它填入下面的 `RUN`：

```sh
RUN='build/runs/startup-communication/startup-实际运行标识'
python3 scripts/debug/startup_communication.py start --run "$RUN"
python3 scripts/debug/startup_communication.py status --run "$RUN"
python3 scripts/debug/startup_communication.py collect --run "$RUN"
python3 scripts/debug/startup_communication.py report --run "$RUN"
python3 scripts/debug/startup_communication.py report --run "$RUN" --detail
```

`start` 上传测试包，校验压缩包 SHA-256，并注册一次性交互式计划任务，触发后立即返回。Windows 自主执行；Mac 退出或断线不影响它。`status` 只读查询，不再次提交、不自动重试。若启动命令断线或超时，先查询状态；同一运行标识禁止再启动。

可用 `--host` 指定其他 SSH alias，`--python 'C:/.../python.exe'` 指定 Windows Python。默认先发现 WorkBuddy 已有解释器，再查 PATH；最终解释器与版本记在证据里，不安装任何依赖。

```sh
python3 scripts/debug/startup_communication.py stop --run "$RUN"
python3 scripts/debug/startup_communication.py status --run "$RUN" --json
```

停止为协作式：当前用例完成或达到看门狗时限后不再启动下一用例。只杀死超时的测试子进程，由其 Job Object 回收所属 bridge；不会杀 WPS，也不会清除文档隔离状态。可能已有文档效果的超时标为 `unknown`。

## 测试文件与前置条件

复制并编辑 [quick.json](../../../src/test/resources/diagnostics/startup-communication/quick.json)，然后 `pack --config 路径 --profile quick`。长时间测试用 `--profile long`，必须显式启动。

- 快速检查默认最多 10 轮、300 秒；长时间默认 7200 秒。预算控制是否开始下一用例；已经开始的用例最多继续 `caseSeconds`（默认 180 秒）。外层计划任务另有 `预算 + 用例上限 + 120 秒` 的兜底上限。
- `cases` 可选 `pipe`、`wps`、`wps-open`。`wps-open` 要求同轮列表前面有 `wps`：校验前一用例产物哈希，复制为本用例输入，重新打开并修改，再独立验证新产物。每轮按列表顺序执行。非预期失败立即停止整次运行；剩余部分未执行，不计作失败。
- `startupSeconds` 是 bridge 就绪握手时限；`requestSeconds` 是纯通信消息时限。WPS Action 仍使用生产默认 60 秒，没有通过增大其超时宣称稳定。
- 所有进程及 Task 都重新建立；每次使用新目录、新提交身份、新文档名。WPS 应用可能已经运行，不宣称测到了 WPS 冷启动。前置条件与执行证据分别留存。
- `assertions` 明确要求消息完全一致、Task 与清理成功，并指定最终 XLSX 的 A1/B1 值。配置不接受静默忽略的字段或无效时限。

第一层用纯 PowerShell 应答夹具，复用生产 Python transport，检查不同长度、中文、Unicode 和转义字符的消息。第二层生成完整 Task JSON，走生产校验、调度、bridge、WPS 创建/读写/保存/回读路径，并独立解析 XLSX 验证内容。第二层不是纯通信性能测试。

目前夹具使用 Excel 承载真实文档操作，不代表 Word/PPT 已验证。简单单元格内容用例不要求视觉检查；配置可声明 `visualReview: required`，届时结果明确标为 `pending`，不能视作人工审阅完成。视觉检查的录入工作流不在本里程碑内。

## 阶段与证据

现有基线链路：

```text
Mac pack → SSH/scp → 交互式计划任务 → 外层 PowerShell
→ Python supervisor → Python Task CLI 测试入口
→ Task-owned PowerShell bridge → WPS COM → 回读/保存验证
```

三个可分别判定的阶段：

1. `bridge.process_create`：创建、归属并恢复 bridge 进程。
2. `bridge.ready_handshake`：应用脚本已加载，可正确应答一个测试专用请求，不调用 COM。
3. `wps.task` + `wps.artifact_verify`：真实 Task 成功，且产物内容独立校验通过。

还记录 Mac 打包/上传/触发/下载、Python 创建、导入、每次 bridge 请求、退出与清理。已有 production trace 保留 JSON 读取/解析、Action 校验/执行、bridge encode/write/wait/decode、原生操作计时。

**生产源码未修改。** 构建时在测试包副本中为共享 bridge loop 插入 `__diagnostic_ready` 分支，并为 Excel COM 注册检查、连接、激活及文档绑定加入诊断包装；Python 测试入口包装启动器和 transport。原有连接回退策略和异常继续保留。变化、修改前后哈希写在 `instrumentation.json`，生产包来源与实际测试文件哈希写在 `manifest.json`。这会增加一次往返、注册表只读查询、WPS 版本读取和阶段日志；以后候选方案必须使用相同测量条件，不能拿未插桩历史数据直接比较。

进程创建与就绪是不同测点。`bridge.wait_response` 包含原生执行，不能当成纯通信耗时。阶段可能嵌套，不应相加。不同主机的单调时钟不得相减；跨机 UTC 只作时间线线索。本里程碑记录计划任务触发与 wrapper 开始时间，但尚未精确拆分任务调度排队、解释器启动的各项贡献。

运行目录包括：

```text
run.json                       Mac 提交身份及远端位置
bundle/manifest.json           commit、dirty 状态、源码与测试包哈希
host-events.jsonl              Mac 阶段耗时
host-logs/                     SSH/scp 原始输出
snapshots/<id>/windows/
  status.json                  逐用例结果、当前步骤、停止原因
  events.jsonl                 Windows supervisor 阶段
  environment.json             解释器和交互会话信息
  cases/r00001-pipe/            消息、启动及请求阶段、清理证据
  cases/r00001-wps/             输入、回执、trace、XLSX、独立检查
```

运行中的 `collect` 生成标为 partial 的快照；读取使用 `FileShare.ReadWrite | FileShare.Delete`，不阻止日志追加或原子替换，并只复制文件开始读取时的长度。正在创建、替换或独占的文件可能未进入部分快照，遗漏会记录在 `snapshot.json`，不把部分快照当完整证据。已完成快照核对证据清单和整个 ZIP 哈希。外层 wrapper 退出时仍会写 `launcher.json`、`worker.stdout`、`worker.stderr`，这三项由快照 ZIP 哈希覆盖，不纳入 supervisor 提前生成的文件哈希清单。每次取回生成新目录，不覆盖旧记录。

成功后只关闭已保存、已独立校验且路径精确匹配的测试文档，保留产物和日志。失败保留仍然存在的文档现场，不擅自保存或关闭；原执行器照常清理其资源。计划任务及远端证据目录保留用于检查，当前没有自动删除命令。环境/原始日志可能含本机路径，不应直接当作脱敏公开报告。

## 验证与后续边界

```sh
python3 -m unittest discover -s src/test/python/tests/diagnostics -v
```

Windows 文件共享回归（仅临时文件，不调用 WPS）：

```powershell
python src/test/python/tests/diagnostics/snapshot_sharing.py --copier src/test/resources/diagnostics/startup-communication/snapshot_io.ps1
```

本地测试检查失败即停、不重放、超时、协作停止、哈希篡改与独立产物验证等行为；不计作 WPS 实机通过率。实机验证结果另见本目录的 `RESULTS.md`。

里程碑 2 比较现有与简化启动链，固定通信；里程碑 3 比较匿名管道、命名管道、TCP，固定启动方式，并增加故障注入矩阵。每个实验声明稳定性、正确性和成本的判断标准。当前没有方案优劣结论，也没有长期可靠性保证。

## COM 环境验收

在 Mac 运行：

```sh
python3 scripts/debug/startup_communication.py pack --config src/test/resources/diagnostics/startup-communication/com.json
```

随后使用上面的 `start/status/collect/report` 命令。这份配置每次跑一轮创建和重新打开，共两个真实 WPS 用例；换机器时用 `start --host 新机器SSH别名 --python Windows解释器路径`。目标用户应已登录桌面，安装 WPS 表格，且能够注册自己的交互式计划任务。

报告逐阶段显示：环境 → COM 注册可发现 → 获取已有实例／COM 激活 → 创建／打开并绑定文档 → 写入 → 回读 → 独立检查 XLSX。记录 PowerShell 版本、位数、权限、会话、线程 apartment、WPS 自报版本，以及 HKCU/HKLM 的 32/64 位注册视图。注册明细是观察值；真正可用性仍以后续真实调用判断。

每个 WPS 用例目录新增 `com/*.jsonl` 与 `com-summary.json`。错误保留异常链、HRESULT 和原始消息；开始后未结束为 `unknown`，未走到的分支为 `not_executed`。获取已有实例失败而激活成功时，两次结果均保留，最终连接路径明确为 `activation`。激活成功不等于已证明冷启动。必要阶段缺少证据也会使验收失败。

要覆盖激活路径，请在专用测试桌面自行保存并关闭 WPS 后，用新的运行目录再测；工具不会为制造该前置条件关闭你的 WPS。一次验收只证明这台机器、这个用户与当时桌面状态下的路径。Word/PPT、其他账户权限组合与其他机器需要分别验收。

Windows 原生负向诊断测试（输出目录必须不存在，仅使用不存在的 COM 标识和注入异常，不操作 WPS 文档）：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File src/test/resources/diagnostics/startup-communication/com_diagnostics_test.ps1 -Helper src/test/resources/diagnostics/startup-communication/com_diagnostics.ps1 -OutputDirectory C:/temp/wps-com-negative-unique
```

当前实测与未覆盖项见 [COM_RESULTS.md](COM_RESULTS.md)。
