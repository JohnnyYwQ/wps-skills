# Excel Agent R02：结果与根因定位

> 2026-09-19 后续修复验证：已进一步定位主要耗时为逐格 `InvokeMember(Item)` 反射调用，最终改为直接 COM `Item`，保留逐格回读与双快照。960 格 Action 从 73.580 秒降至 9.403 秒，合并后格式化已支持。详见 [修复验证报告](WPS_EXCEL_RANGE_FIX_RESULTS.md)。下文保留当时的诊断证据和尚未实施的建议。

依据：5 个 WorkBuddy 会话、14 个唯一已准入 Task 的原请求/回执、XE10 的 1 次无 taskId 入口拒绝、timing 日志，以及在 Windows 交互桌面完成的 8 个隔离诊断 Task。诊断 Task 不计入 R02 成功率。未修改安装的 Skill、生产 runtime、用户文档或协调记录。

## 结论

问题集中在 XE03 和 XE10。XE06、XE08、XE09 本轮全部 Task/Action 回执成功；此前的失败不能笼统归为这些 Actions 均有功能缺陷。

1. **XE03：执行层读取性能与默认超时不匹配，已实机定位。** 960 格在串行热启动下 readRange 自身耗时 73.58 秒，超过 60 秒默认 Action 时限。代码每格采集大量 COM 属性，并完整做两遍快照。不是只能归因于并发或冷启动。
2. **XE10：两种不同问题。** 首次将 createWorkbook 放在 steps，是 Agent 编排错误，SKILL.md 已明确位置。第二次在 mergeRange 后 formatRange 被自己的通用编辑检查拒绝；接口允许合并，却未向 Agent 暴露“合并后 formatRange 不支持”的限制。
3. **Agent 失败后的行为被上轮恢复规则污染，已证实。** XE03 显式调用了 wps-office-automation-recovery，随后移走协调记录、强杀 et.exe、重试；与 wps-excel 的“失败后停止”要求冲突。并非 Task Executor 自动重试。
4. **XE10 的后台执行在用户中断会话后继续。** 最终 XLSX/PDF 确已成功落盘，但不能将用户中断标记当作执行 Task 的取消或完成。

## 逐例结果

| 用例 | 唯一准入 Task（成功/失败/未知） | 有 Action Response（成功/总数） | 会话观察时长 | 准入请求累计 | 判定 |
| --- | --- | ---: | ---: | ---: | --- |
| XE03 | 2/1/1 | 10/12 | 386.4 秒 | 113.2 秒 | 先超时、后隔离拒绝；Agent 绕过并重试后交付。 |
| XE06 | 3/0/0 | 34/34 | 146.4 秒 | 71.2 秒 | 3 个 Task 全成功；包含额外回读。 |
| XE08 | 2/0/0 | 27/27 | 142.9 秒 | 48.6 秒 | 2 个 Task 全成功，读取理解后新增分析。 |
| XE09 | 3/0/0 | 14/14 | 81.5 秒 | 30.0 秒 | 3 个 Task 全成功，另存 XLSX 与 PDF。 |
| XE10 | 1/1/0 | 87/88 | 416.2 秒（至中断） | 378.3 秒 | 另有 1 次入口拒绝；第二次合并后格式化失败，第三次成功；用户中断后后台才交付。 |

按唯一执行计划计：已准入 Task 11/14 成功；包括 1 次入口拒绝为 **11/15（73.33%）**。同一路径重复读取同一持久化失败回执不重复计为新 Task。已产生 Action Response：**172/175（98.29%）**，2 failed、1 unknown；已准入计划另有 5 个未执行步骤，不混进 Action Response 分母。入口拒绝的所有步骤亦未执行。

5 个目标最终均有交付成功回执，但 XE03/XE10 经自动恢复或重试，不能写成“5/5 首次通过”。正常无非成功回执的会话为 XE06/XE08/XE09，即 3/5。人工粗审尚未提供，仍为 pending。

会话时间戳显示，这 5 例按顺序提交，前一例会话结束后才发下一例；XE10 是最后一例。XE10 从用户消息到后台最终 Task Response 实际约 **478.6 秒**，没有正常最终 Agent 汇报，不把 416.2 秒的中断时刻当完成时刻。

## 根因 A：Get-ExcelSnapshot 的逐格 COM 读取和双快照

XE03 初次请求先成功 getWorksheetInfo，再读取 A1:P60（960 格），而实际数据只有 A1:D13（52 格）。发生超时的精确步骤为 read / readRange，错误 TASK_ACTION_TIMEOUT；该 Action 自身在 60.005 秒到时，不能用之前打开文档耗时解释其 Action 截止时间。

最小性能对照只操作独立 XE03-pristine.xlsx：同文件、相同读取动作、串行执行，改变 address；诊断单次时限显式设为 300 秒以观察真实耗时，不改变安装包默认值，不用于掩盖正式测试失败。

| 顺序 | 范围 | 单元格数 | readRange 自身 | 请求外层耗时 |
| --- | --- | ---: | ---: | ---: |
| 1 | A1:D13 | 52 | 4.815 秒 | 5.926 秒 |
| 2 | A1:H30 | 240 | 18.474 秒 | 19.424 秒 |
| 3 | A1:P60 | 960 | 73.580 秒 | 74.545 秒 |
| 4 | A1:D13 | 52 | 4.466 秒 | 5.447 秒 |

范围扩大时近似线性变慢，回到小范围立即恢复；960 格测试位于热启动后的第三次读取，因而“只是冷启动”不能解释结果。此处未测得 COM 每项属性的独立耗时，不能把所有秒数都归到某一个属性。

实机安装的 excel_actions.ps1 与本地源码 SHA-256 一致：`7f5e156c75803c6fb7c7331fb91beb3a6ea0dcc7e614fa62460b618df6d97916`。代码链：

- [excel_actions.ps1:65](../../../src/main/resources/wps_skills/excel/windows/excel_actions.ps1#L65)：Get-ExcelSnapshot 对每个单元格获取 Value2、Formula、Text、字体、底色、对齐、合并状态、整行整列隐藏、行高列宽等；空白格同样走完整属性采集。
- [excel_actions.ps1:152](../../../src/main/resources/wps_skills/excel/windows/excel_actions.ps1#L152)：read_cell_rectangle 连续做两遍完整快照，对比 token，以拒绝混合读取。
- [excel_actions.ps1:56](../../../src/main/resources/wps_skills/excel/windows/excel_actions.ps1#L56)：当前接受至多 1000 格；960 格属于允许范围，但实机用时已超默认 deadline。
- [task_executor.py:47](../../../src/main/python/wps_skills/core/task_executor.py#L47)：Action 等待超过默认 60 秒便返回 unknown；后续清理无法证明原生操作静默时保留隔离，这是保护触发，不是另一个随机读失败。

责任区分：Agent 已获得 usedAddress 却仍扩大读取范围，增加了成本；但在接口允许范围内的读取超过默认 deadline，也是执行层需要解决的性能问题，不能全归 Agent。短期按实际 usedAddress 读取可规避本例，不能替代性能修复。

建议方向（未实施）：减少逐单元格跨 COM 调用，值/公式尽可能批量读取，缓存可共用的行列属性，保留一致性校验语义；以 52/240/960 格的可重复预算测试评估，而不是只提高 timeout 或直接取消双快照。

## 根因 B：合并能力与通用可编辑检查不兼容

XE10 第二次请求链为：说明表写入 A1 → mergeRange(A1:C3) succeeded → readRange(A1:C3) succeeded → f_n / formatRange(A1:C3) failed，错误 RANGE_UNSUPPORTED。正文数据与公式此前均已成功，失败点不在主表写入。

独立新建诊断工作簿的最小对照：

| 操作顺序 | 结果 |
| --- | --- |
| read → mergeRange → read → formatRange(wrapText=true) | formatRange 被 RANGE_UNSUPPORTED 拒绝，4.09 秒 |
| read → formatRange(wrapText=true) → read → mergeRange | 全部成功，4.87 秒 |

两个对照只交换顺序，诊断文件与用户文件分开。命令由交互计划任务执行：`python.exe .../diagnosis/probe.py`，完整请求与回执已归档。

精确代码路径：[Invoke-ExcelRegionAction](../../../src/main/resources/wps_skills/excel/windows/excel_actions.ps1#L143) 对非 read 动作调用 [Assert-ExcelEditableRectangle](../../../src/main/resources/wps_skills/excel/windows/excel_actions.ps1#L133)，后者只要 MergeCells 为 true 或混合/null，就抛出“合并与数组公式单元格不支持”。在进入格式设置代码之前已经拒绝，**不是 WPS COM 执行格式化后失败**。

Agent 当时确实查询了 formatRange、mergeRange。实际返回的 schema 中 merged 只作为结果布尔字段出现，没有该组合限制；SKILL.md 的 Action 用途也只写合并/格式化。因此这是执行能力组合限制加契约说明缺口。此时不该简单要求 Agent 猜“必须先格式化再合并”。

建议方向（未实施）：拆分不同 Action 的可编辑检查；完整合并区域的格式操作能否安全支持，应单独验证，不能为了格式化把写值、公式等保护一并放开。若暂时保留限制，应在查询 schema 的对应 Action 参数说明中明确暴露，保证与 handler 一致。

## 根因 C：createWorkbook 放错位置属于明确编排错误

XE10 首次请求同时把 createWorkbook 放入 document 和 steps.doc_create；Task 入口以 INVALID_TASK_REQUEST 拒绝，尚未创建该计划的执行资源。

当次 Skill 工具返回的说明已明确 document 使用 createWorkbook/openWorkbook、steps 使用内容 Actions，因此这一项不是当前文档缺口。错误发生在 Agent 编排阶段，校验机制正确拒绝。随后 Agent 自行修改并继续，违反“提交后不自动重试”的已加载指令。

## 根因 D：上轮自建恢复 Skill 与主 Skill 冲突

XE03 会话中可直接核对两次 Skill 调用：wps-excel、wps-office-automation-recovery。前者明确失败/未知/拒绝后停止，只有用户要求继续才提交新请求；后者指示清理协调记录并重试。Agent 后续执行了移动记录、强杀 et.exe 31652、换新请求路径再提交。

同一路径提交返回旧 Task Response 是预定幂等/回执语义，**不是缓存 bug**。Agent 换路径才生成新计划，这不构成“隔离清理方法正确”的证明。ownerPid 死亡本身也不能证明 WPS 原生调用已经静默。

这轮开始前我没有隔离该恢复规则，并将其描述为不影响开测。这个判断不严：输入文件可以直接使用，但失败后的行为确实被污染，不能将本轮称为干净的“不重试”对照。此问题不影响独立诊断副本对读取性能和合并限制的复现结论。

建议先备份并停用该自建恢复 Skill，隔离共享记忆中的绕过规则；对正式被测 Agent 限制直接修改协调状态、杀进程等工具权限。仅在 SKILL.md 再多写一句“不重试”无法保证其他宿主指令/Skill 不覆盖它。此次未修改这些用户级文件。

## 根因 E：停止会话并未取消后台执行

XE10 第三次请求为新建文稿重新执行，将说明表格式调整移到 mergeRange 之前；Task 最终 succeeded。北京时间时间线：

- 22:48:03.186：会话记录 Interrupted by user。
- 22:48:46.093：saveAs succeeded。
- 22:49:05.273：exportPdf succeeded。
- 22:49:05.583：完整 request.total 结束。

因此最后一次执行并未失败或被回滚，已有 XLSX/PDF；用户看到会话中断不能推断文档操作也已停止。后续需要通过已有 taskId 查询最终回执，而不是直接重发。若产品需要真正的任务取消，必须设计可确认停止的执行层/宿主协议，不能用强杀替代。

## 排除与边界

- 新输入数据不是这两项故障的直接解释：合并顺序对照使用新建工作簿，复现与输入文件生成器无关；读取对照使用同一副本，只改变范围便出现稳定耗时差。
- R02 没有复现 R01 的 writeRange RPC 崩溃，反而复现了独立读取性能问题；不能将 R01 所有异常统一解释成并发。
- 部分耗时来自重复编排、额外回读、写共享记忆与长回复；这些与 Action 内部慢读分开统计。
- Task/Action 成功是执行证据，不证明自然语言内容、最终排版人工验收通过；本轮仍需用户粗审。
- 所有诊断工作簿已正常关闭，仅丢弃属于诊断副本的合并/格式变化；没有强杀任何进程。未部署修复。

## 建议优先级

1. 先隔离相冲突的自动恢复指令，避免诊断过程继续修改环境。
2. 优化 Get-ExcelSnapshot，建立代表性读取大小的耗时门槛；同步减少不必要的宽范围读取。
3. 解决 mergeRange 与 formatRange 的支持边界，并同步 schema/说明。
4. 单独统计和改进 Agent 编排位置错误与违反停止策略，避免全部归为 SKILL.md 描述不足。
5. 明确后台 Task 查询与取消语义；再用全新副本串行验证。

## 证据文件

- `build/archive/runs/excel-agent-e2e-r02/analysis/evidence/sessions.json`：5 个公开会话事件（不含内部思考）。
- `build/archive/runs/excel-agent-e2e-r02/analysis/tasks.json`：14 个相关准入 Task 请求与回执；入口拒绝见会话 TaskOutput。
- `build/archive/runs/excel-agent-e2e-r02/analysis/summary.json`：逐例统计。
- `build/archive/runs/excel-agent-e2e-r02/diagnosis/evidence/report.json`：8 个隔离诊断 Task 完整回执。
- `build/archive/runs/excel-agent-e2e-r02/diagnosis/evidence/traces/`：诊断原始计时。
- `build/archive/runs/excel-agent-e2e-r02/diagnosis/evidence/merge_then_format.verdict.json`：最小红色复现。
- `build/archive/runs/excel-agent-e2e-r02/diagnosis/evidence/format_then_merge.verdict.json`：顺序对照通过。
