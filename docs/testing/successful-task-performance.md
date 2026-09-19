# 正常成功任务的执行耗时

2026-09-19，仅离线复算全 Action 验收 v6 的已有回执与 trace，没有重新执行实验。选择这一统一批次，避免混合更早版本、真实 Agent 编排耗时及设计消融候选。

## 统计范围

纳入 manifest 预先标记为 positive、Task succeeded 且验收断言通过的业务用例。29 个正常成功 Task 中，X10-setup、P10-setup 是准备已有文档的夹具，另列并排除；最终性能样本为 **27 个业务 Task、639 次 Action**。

另外排除 85 个入口拒绝、X12/P12 两个预期失败 Task（连同它们已成功的前序 Action）。不纳入今日设计实验的故障注入、消融、重复只读成本样本。成功条件是固定请求执行与验收断言，不等于全部自然语言任务已人工验收。

## 结果

| 应用 | Task 样本数 | Action 次数 | Task 累计 | Task 平均 | Task 中位数 | Task 范围 | 单 Action 中位数 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Word | 3 | 21 | 6.65 s | 2.22 s | **2.16 s** | 1.69–2.80 s | **174.82 ms** |
| Excel | 12 | 250 | 147.87 s | 12.32 s | **8.92 s** | 3.03–47.13 s | **167.14 ms** |
| PPT | 12 | 368 | 49.78 s | 4.15 s | **3.73 s** | 1.68–7.00 s | **53.51 ms** |
| 合计 | **27** | **639** | **204.31 s** | 7.57 s | 4.59 s | 1.68–47.13 s | 75.03 ms |

各应用用例工作量不同，每个 Task 的 Action 总数为 Word 4–10、Excel 10–41、PPT 9–75。整体平均/中位数受用例配比影响，对外优先按应用分别报告。Excel 较长的成功任务仍保留，没有根据耗时删掉慢样本。

Action 次数包含新建/打开、业务步骤及保存/导出。仅内容步骤共 577 次；其 Action 中位数 Word 172.83 ms、Excel 116.71 ms、PPT 51.14 ms。如使用该子集，必须写明“不含文档获取和交付步骤”，不替换上表全 Action 口径。

## 计时边界与复算方法

- Task 采用原执行请求 trace 的 `request.total`：从 CLI 内部处理输入到响应写出，包含验证、文档获取、执行、回执与资源清理；不包含解释器启动、Agent 思考/编排、用户等待和外部验收。
- Action 采用同一执行请求的 `action.execute`，包含原生执行与内部回读。它被 Task 时间包含，两者不能相加。
- 不使用 `report.json` 的 case.seconds：该值还包含回执查询、产物检查和测试文档关闭，会高估被测执行器的耗时。
- 按唯一 taskId 关联成功回执与 trace，仅选择具有 action.execute 的原执行请求，避免重复计入后续回执查询。27 个 Task 均有且只有一组执行 trace；639 个 Action 计时与对应回执步骤数一致。
- 同一 Windows 主机、WPS 已运行，每用例一次；这是成功任务的执行性能基线，不是冷启动保证，也不是端到端 Agent 耗时或相对于其他设计的提速比例。
- 只选正常成功样本用于条件耗时统计，不据此重算项目成功率。失败、拒绝和历史返工仍在各自结果报告中。

## 简历可用句子

> 设计任务级批量执行与资源管理机制，统一处理步骤依赖并在任务内复用执行资源；在包含读写、格式设置和保存导出的实机验收中，Word、Excel、PPT 成功任务的执行耗时中位数分别为 2.16、8.92、3.73 秒。

## 证据

- [汇总、筛选条件与源文件 SHA-256](../../build/evidence/action-acceptance-v6-performance/summary.json)
- [逐 Task 耗时](../../build/evidence/action-acceptance-v6-performance/tasks.csv)
- [639 次 Action 耗时](../../build/evidence/action-acceptance-v6-performance/actions.csv)
- [按 Action 名称分组](../../build/evidence/action-acceptance-v6-performance/per-action.csv)
- [原全 Action 验收报告](action-execution-results.md)
