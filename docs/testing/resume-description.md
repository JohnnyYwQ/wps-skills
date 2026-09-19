# 简历表述：WPS Agent Plugin

按当前事实组织为“达成结果—设计贡献—实测证据”。不以已修复的 Excel 性能 bug 作为设计优化亮点，不纳入本次向前查到的旧版 Word 批次，不写具体 Agent 宿主名称。

## 可用描述

**WPS Agent Plugin｜面向智能体的本地办公自动化插件**  
Python、PowerShell、Windows COM、JSON Schema

- **实现真实文档操作能力**：为 Word、Excel、PPT 封装 85 种业务 Action，支持 Agent 从自然语言需求生成执行计划并完成本地文档编辑与保存导出；完成 Word 20、Excel 10、PPT 10 个场景的真实 Agent 测试，含定向复测累计产生 113 个已受理 Task，另以固定请求覆盖全部 85 种 Action 的成功执行与响应断言。
- **通过渐进式披露降低上下文成本**：采用“用途指引选择操作→按需查询完整 schema”的流程，避免每次加载全部操作定义；在每应用固定选择四个 Action 的对照中，schema token 数按参考分词器统计减少约 74%–86%。
- **通过 Task 编排减少逐步交互与重复初始化**：设计任务级批量执行与资源管理机制，统一处理步骤依赖并在任务内复用执行资源；在包含读写、格式设置和保存导出的实机验收中，Word、Excel、PPT 成功任务的执行耗时中位数分别为 **2.16、8.92、3.73 秒**。
- **防止执行错误扩大并保留已知事实**：通过整计划预校验、精确绑定、原生回读、提交身份与逐步持久回执，在真实 WPS 对照实验中验证末尾参数错误不造成部分修改、切换活动窗口不误改其他文档、漏写不会假报成功，以及原路径重复提交不重放、执行中断后已完成保存事实仍可查询。

## 真实 Agent 指标的选用口径

Word 的主体为 20 个场景，E03/E10/E12/E19 是其中四个场景修复后的复测，不是全部 Word 测试只有四个。首轮产生 32 个受理 Task，复测新增 4 个；累计 36 个 Task 中 32 succeeded、2 failed、2 unknown。四例复测耗时中位数 35.33 秒只能描述该复测子集，不能作为全部 Word 场景的耗时。

结合 Excel/PPT R01 的 63 个 Task、Excel R02 的 14 个 Task，按唯一 taskId 复核共 **113 个受理 Task：95 succeeded、10 failed、8 unknown**；另有两次无 taskId 的入口拒绝。40 个独立场景加 9 次定向复测，共 49 次场景运行。历史 Word 路径前置条件的限制、失败后返工和保护绕过仍保留在报告中；这些数量是实际测试规模，不是统一最终版本的完成率。

今天全 Action 固定请求批次的 **29/29** 只指其中正常成功的 Task；同批另有 2 个预期失败 Task、85 个入口拒绝。它既不是全部项目 Task 数，也不是 Agent Task 数。按 Task 重新核对的明细见 [计数审计](../../build/evidence/agent-design-synthesis-20260919/task-count-audit.json)。

## 数字与归因

- 渐进式披露的 token 以已有 Windows stdout 为输入，tiktoken 0.14.0、`o200k_base`：Word 9327→2394，Excel 6563→1139，PPT 8344→1144；按需均为四个 Action。`cl100k_base` 的百分比接近。它衡量 schema 载荷，不能写成总会话或实际计费 token 减少 74%–86%。
- 简历使用全 Action 验收 v6 的正常成功业务样本：27 个 Task（Word 3、Excel/PPT 各 12）、639 次 Action；排除两条素材准备 Task、参数拒绝及预期失败任务。Task 用 request.total，包含 CLI 内部输入处理、文档获取、执行、回执与清理，不含解释器启动或 Agent 编排。这与此前三应用 × 三种只读长度 × 三次重复的 27 个成本样本是两批不同数据。详见 [正常任务耗时](successful-task-performance.md)。
- 成功任务耗时体现已达到的执行性能，不是设计相对于逐 Action 调用的加速百分比，也不据此筛选后计算成功率。
- bridge 的任务内复用及结束后释放另有 [生命周期核对](../../build/evidence/design-value/bridge-lifecycle-audit.json)，没有跨 Task 常驻复用。bridge 释放不关闭 WPS 应用或用户文档；持久回执的保留期也不等于进程生命周期。
- 保护设计的效果均限定于已测场景；精确绑定和回读不能替 Agent 判断它是否在请求中选错目标页或填错业务内容。
- 19/20、4/4 是历史指定批次的实际表现；不将今天固定 JSON 的成功追加到 Agent 完成率分母，也不拿修复后的结果覆盖历史失败。

证据：[执行端验收](action-execution-results.md)、[机制与成本实验](host-independent-results.md)、[Agent 批次与计时](agent-and-design-evidence.md)、[token 复算及源文件哈希](../../build/evidence/design-value/schema-token-counts.json)。本次只复算既有数据，没有启动新测试。
