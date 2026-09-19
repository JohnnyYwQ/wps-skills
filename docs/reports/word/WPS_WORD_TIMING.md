# Word 正式执行链路耗时日志

已在正式 Word Task 调用链记录耗时，不依赖测试入口 monkey patch。Action/Task Response 结构和 contract 均保持不变；文字测试后续直接使用正式 trace。图片尺寸及分页计数问题未改动。

## 日志位置与关联

默认 Windows 目录：`%LOCALAPPDATA%/wps-skills/logs`；可通过进程环境变量 `WPS_TRACE_DIR` 指定。日志为 JSONL。

- `requests/request-*.jsonl`：从 CLI 开始处理 Task 输入到响应写出/flush 的整条 Python 执行链。
- `bridges/bridge-<pid>.jsonl`：Word PowerShell 端的解析、执行和响应写出耗时。
- 既有 `tasks/`、`actions/` trace 保留。

每次 CLI 调用使用独立 requestId，包括非法 JSON、回执查询和重复提交；Task 受理后通过 `request.task_bound` 关联 taskId。受理前记录的 taskId 为 null，应按同一 requestId 关联，不能把它算成日志缺失。查询旧回执只记录读取与返回，不伪造 Action 耗时。

Python span 带 spanId/parentSpanId、UTC时间、单调时钟起止值、durationNs/durationMs，以及适用的 taskId、stepId、action、traceId、bridgeRequestId。Action结束记录outcome；抛出异常记录exceptionType，日志不记录请求正文、参数或文档内容。`status=returned` 只表示调用正常返回，不代表业务成功，业务结果看 outcome。进程被强杀时可能只有开始记录，不能补零。

跨线程传递计时上下文，包括Action执行、文档资源释放和adapter清理。PowerShell通过bridgeRequestId/traceId与Python对应；不同进程的时钟起点不能直接相减。

## 各时间的含义

| 事件名 | 边界 |
| --- | --- |
| request.total | CLI参数解析和日志初始化完成、准备接收Task文件 → Task Response编码、stdout写入及flush结束，含中间清理 |
| input.read / input.decode | 输入文件读取；JSON解析 |
| task.client | Task Client处理请求，含预检、受理、执行、回执和清理，不含CLI最终输出 |
| task.validate_request / task.preflight | 请求结构校验；Action参数和引用静态预检 |
| task.admission_lock / task.admit_request / task.input_consumption | 受理锁、原请求落盘、受理后输入文件消费 |
| task.execution / task.resources_startup | 已受理Task的执行与清理总区间；运行资源组装 |
| action.resolve / action.validate | 逐步引用解析、Action参数校验 |
| action.execute | executor.execute调用 → 结果返回，包括线程等待、桥接通信、WPS操作、回读与结果处理 |
| action.runtime | Action工作线程中的共享执行机制；关联同一step和trace |
| bridge.process_start | 创建PowerShell桥接进程；不等于桥接已完成启动 |
| bridge.round_trip | 一次桥接exchange入口 → JSON响应解析完成；包含排队、通信与远端执行 |
| bridge.encode / bridge.write_flush | 桥接请求JSON编码；管道发送与flush |
| bridge.wait_response | 等待响应完整行，包括WPS执行、回读和响应传回，**不是纯通信耗时** |
| bridge.decode | 响应JSON解析 |
| native.decode_validate | PowerShell收到完整行后解析和验证请求 |
| native.operation | PowerShell Invoke-BridgeOperation，包含协调、WPS COM调用和内部回读，不是单个COM方法耗时 |
| native.response_write | PowerShell响应编码、输出（使用既有Write-BridgeRecord） |
| task.receipt_publish / response.progress_write | 回执落盘；进度写出 |
| task.document_check | 获取文档后检查Task要求的前置状态 |
| task.cleanup | 文档资源、桥接及自有进程释放 |
| response.encode / response.write_flush | 最终Task Response编码；stdout写出和flush |

`request.total` 不包含调用方生成/传输JSON到磁盘、Python解释器及模块加载、调用方收到stdout后的处理。进程内部无法测量这些外部环节；测试仍可额外测量启动CLI到接收完成的端到端耗时，并用不同名称报告。

不能把 `bridge.round_trip`、`action.execute`、`task.execution` 和 `request.total` 相加：它们存在包含关系。Task内各 `action.execute` 可相加；Task总耗时减Action合计是其他开销，不应命名为纯通信。桥接RTT减PowerShell执行时间仍包含启动、排队、序列化、系统调度和日志成本，仅可视作剩余开销估算；当前不输出虚假的“纯通信”数字。

## 日志可靠性

耗时使用Python perf_counter_ns和PowerShell Stopwatch计算，避免系统时间校准造成负时长。UTC仅用于定位时间。日志写入失败按best-effort处理，不改变Action结果或触发重试；异常和超时保留已记录区间。日志同步写入有小幅开销，统计值是启用日志的真实运行耗时，不宣称零开销。

## 验证

- 本地全套311项：309通过，2项Windows专用测试跳过。
- 补齐清理线程上下文后，相关55项回归通过。
- 新增测试覆盖单调时钟、嵌套span、跨线程关联、坏JSON到响应写出的全链路、日志不可写不掩盖业务异常、stdout断开时的终态记录。
- Windows实机测试使用新建、四次不同文字格式写入、另存，共6个Action。最终包与实测结果见本文件后续记录。

## 最终 Windows 验证结果

最终独立包：`build/archive/runs/word-timing-final-20260917/wps-word`。包manifest校验通过，actions.json与指纹修复基线逐字节相同。

2026-09-17交互桌面实测：Task成功，6个Action全部成功；15次Python桥接往返与45条PowerShell阶段记录完全对应（每次3阶段），包括清理时的释放请求。CLI内部request.total=2182.195 ms，外部提交进程端到端=2267.7755 ms。createDocument=1268.1282 ms；四次writeContent依次274.2131、103.2198、71.5071、73.7923 ms；saveAs=181.5973 ms。该数据是单次功能验证，不是性能基准。

[实机报告](../../../build/evidence/word-timing-20260917-233805/windows/report.json)、[完整日志](../../../build/evidence/word-timing-20260917-233805/windows/logs)、[原始Task Response](../../../build/evidence/word-timing-20260917-233805/windows/stdout.json)。前一轮验证发现清理线程缺计时关联，已修正并以新的独立Task重新验证；未重放旧Task。全部临时计划任务已清理，文档和证据保留。

文字27–38的全套测试尚未执行；本次使用35的文字组合进行计时链路验证。


三阶段全量验证：38个Task、226个Action、566次桥接通信均有完整计时，见 [统一结果](WPS_WORD_THREE_STAGE_RESULTS.md)。分析时须用bridgeRequestId与operation组合匹配：同一Action可能复用请求ID执行多个不同operation，不能把它们的原生耗时混为一次往返。
