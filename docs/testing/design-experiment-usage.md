# 设计实验的终端用法

[实验判断标准](host-independent-experiments.md) · [结果与结论边界](host-independent-results.md) · [早期失败台账](design-experiment-failures.md)

测试代码在 `src/test/python/diagnostics/design_value/`，PowerShell 观察/注入夹具在 `src/test/resources/diagnostics/design-value/`。生产源码与测试候选分别打包；补丁、原文件哈希、修改后哈希见包内 `instrumentation.json`。测试包不是供宿主安装的正式 Skill。

已交付 [Windows 独立实验包](../../build/distributions/wps-design-value.zip)，原样复制自实际运行的 v13；历史实验各自的包版本见结果页。本轮快速实验已结束，通信长测已按要求停止。以下命令供以后手工使用，本轮不再启动或复测。

## Mac 统一启动、回连、取证

要求 SSH 别名 `win` 指向已登录 Windows 桌面的测试账户，Python/WPS 环境先通过 doctor。启动命令只上传包并触发 Windows 交互计划任务；随后 Mac 可以断开。每次使用新的包目录、运行名和结果目录。

```sh
# 固定当前工作区为独立、带哈希的测试包
python3 scripts/debug/design_value.py pack --output build/tools/design-value-my-run

# 三应用机制；--apps word/excel/ppt 可以只跑一个应用
python3 scripts/debug/design_value.py start \
  --kit build/tools/design-value-my-run --name mechanisms-my-run \
  --groups references preflight replay focus dynamic concurrency stale cross_task response_loss persistence \
  --trials 3

# 随时回连观察；不触发重放
python3 scripts/debug/design_value.py status --run build/runs/design-value/mechanisms-my-run
# 可选只读采样：所属进程的内存、句柄、线程
python3 scripts/debug/design_value.py resources --run build/runs/design-value/mechanisms-my-run

# 结束后取回共享读取快照，逐文件核验 SHA-256
python3 scripts/debug/design_value.py collect --run build/runs/design-value/mechanisms-my-run

# 本地重验哈希并计算汇总，可给多个 --runs 参数值
python3 scripts/debug/design_value.py report \
  --runs build/runs/design-value/mechanisms-my-run \
  --output build/evidence/design-value/my-summary.json
```

后续批次使用新 `--name`，按需选择：

| `--groups` | 应用与建议次数 | 用途 |
| --- | --- | --- |
| `interruption receipts isolation lease_paths` | 三应用，`--trials 1` | 受控终止 Python、超时窗口、Job 进程终态与隔离 |
| `readback readback_cost` | `--apps excel --trials 3` | 漏写故障对照与正常写入的 24 对成本样本 |
| `schema` | 三应用，`--trials 1` | 每应用内置 10 对 schema 查询 |
| `task_cost` | 三应用，`--trials 3` | 每轮 1/8/32 个内容 Action，轮换顺序 |
| `protocol startup` | `--apps word --trials 1` | 公共协议只跑一份；这里的 word 只是去重选择，不调用 Word 或 WPS |

含真实文档的批次顺序运行，避免额外焦点切换影响解释；纯通信可并行，但性能报告必须注明负载条件。

## Windows 手工执行

将测试 ZIP 解压到独立目录，用 doctor 已验证的 Python 在已登录桌面运行：

```powershell
python C:/wps-tests/design-kit/run.py --root C:/wps-tests/runs/design-new-01 `
  --apps word excel ppt --groups references preflight replay focus dynamic concurrency stale cross_task persistence --trials 3
```

`--root` 必须尚不存在。可以分别用 `--apps word`、`--apps excel`、`--apps ppt` 分批运行。不要在 SSH Session 0 直接执行 COM；Mac 入口已通过交互计划任务安排正确桌面。

## 出错后查哪里

运行返回失败会停止该批次，不自动修正 JSON、重新执行旧 Task 或继续其他场景。失败与预期故障是两回事：预期的拒绝、unknown 和隔离通过实验断言后可以继续。

1. `results/report.json`：失败实验与未完成位置。
2. `results/calls/<序号-名称>/request.json`、`config.json`：原始输入和实际注入点。
3. 同目录 `response.json`、`stdout.jsonl`、`stderr.jsonl`、`child-events.jsonl`、`native-events.jsonl`：响应、具体 Action 阶段和桥接事件。
4. `results/observations/` 与 `results/outputs/`：独立 COM 观察和保存产物；不能只依据 Action 的成功字段。
5. `results/receipts/`、`results/traces/`：持久回执与阶段耗时。部分计时包含子阶段，不能全部相加。
6. `collection.json`、`collected/files.json`：归档与逐文件校验结果。

仅已保存并验证的成功文档会按确切路径关闭。失败或未保存的测试现场保留；不会退出 WPS、关闭无关文档、删除 Quarantine 或强制解锁。中断/超时不能授权重放原 Task。需要新实验时，使用新的运行根目录和请求路径。

## 持续通信

独立使用启动通信工具：

```sh
python3 scripts/debug/startup_communication.py pack \
  --config src/test/resources/diagnostics/design-value/pipe-long.json --profile long
# 把上一步打印的新目录传给 start/status/collect/report
python3 scripts/debug/startup_communication.py start --run <新目录>
python3 scripts/debug/startup_communication.py status --run <新目录>
python3 scripts/debug/startup_communication.py collect --run <新目录>
python3 scripts/debug/startup_communication.py report --run <新目录>
```

长测预算 7200 秒，32/4096 字符载荷；每轮新建所属进程、验证应答并关闭。`status.json` 只保存最近 50 轮与累计数，完整历史在 `cases.jsonl`。最终结论同时看监督器、外层启动器和计划任务状态；旧的 `running` 文件不代表进程仍在运行。
