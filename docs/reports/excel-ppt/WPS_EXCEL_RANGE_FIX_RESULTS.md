# Excel 读取性能与合并后格式化修复验证

日期：2026-09-19。结论：两项修复已通过 Windows 原生执行回归，并同步到 Windows 已安装的 wps-excel。此次是直接 JSON 执行侧验证，不计为新增 Agent 端到端成功率样本。

## 修复与精确根因

### readRange

此前已证明 960 单元格读取约 73.58 秒，超过默认 Action 时限 60 秒。本轮进一步分段计时发现，主要热点不是某个内容或格式属性，而是获取每个单元格时调用 `.GetType().InvokeMember('Item', ...)` 的反射开销。

在同一台 Windows、240 个单元格的单次快照探针中，仅该获取步骤累计耗时由 8.611 秒降为 0.098 秒。最终只将快照中的该调用替换为直接 COM `Cells.Item(row, column)`。

原有逐格值、公式、显示文本、格式、错误值、合并与行列状态的读取均保留；两次快照和 token 一致性校验均保留。没有提高默认超时、缩减返回字段或降低读取上限。早期尝试的批量值读取和格式缓存已经撤回，未进入安装版。探针中的分段时间不能直接等同于完整 Action 时间。

### formatRange

原先通用编辑保护拒绝所有合并单元格，导致合法的“合并后设置格式”在进入格式设置前失败。

现在仅对格式化允许完整合并区域：先检查矩形包含所有相交合并区域的完整范围，再对整个请求矩形设置格式，并按原有规则逐格回读验证。局部合并区域仍拒绝；内容写入、数组公式、工作表保护、只读和过期 token 的保护规则没有放宽。

契约中 `formatRange.params.address` 现说明完整合并区域要求及保护限制。对应查询 schema 已重新生成；没有向 SKILL.md 添加 COM 实现细节。

## 最终执行侧结果

隔离根目录：`C:/Users/yim/wps-excel-range-fix-20260919-final2`。交互式 Windows 会话串行执行，使用默认 60 秒 Action 时限。

- 7/7 个 Task 成功。
- 85/85 次 Action 成功，覆盖 33 种 Action，计入建立文档、正文步骤及保存/导出。
- 7/7 Task 资源清理成功。
- 另有原生快照与保护回归通过，共 156 个断言；这不是 156 个独立任务。
- 本地应用层测试 11/11 通过。

| 读取范围 | 单元格数 | 修复前 Action | 修复后 Action | 修复后整次请求 |
| --- | ---: | ---: | ---: | ---: |
| A1:D13 | 52 | 4.815 秒 | 0.905 秒 | 2.008 秒 |
| A1:H30 | 240 | 18.474 秒 | 2.878 秒 | 3.837 秒 |
| A1:P60 | 960 | 73.580 秒 | 9.403 秒 | 10.334 秒 |
| A1:T50 | 1000 | 未测 | 9.811 秒 | 10.743 秒 |

960 格 Action 耗时减少约 87.2%。前后使用同源输入副本，原表实际使用区域为 52 格，扩大的读取范围包含空白格。以上是本机该轮单次测量，不能外推为所有复杂度文档的 P95 或吞吐保证。

其余三个 Task：

| 场景 | 整次请求 | 结果 |
| --- | ---: | --- |
| 合并 A1:C3 后，在同一区域设置九类格式并保存 | 2.408 秒 | 成功 |
| 合并 A1:C3 后，在包含该合并区的 A1:D4 设置九类格式并保存 | 2.652 秒 | 成功 |
| 复用 Excel 综合执行场景，包含读写、格式、查找替换、排序过滤、行列、工作表、公式、保存及 PDF 导出 | 12.901 秒 | 成功 |

`bridge.wait_response` 包含原生操作执行，不能全部计作通信开销。例如 960 格场景 bridge round trip 约 9.349 秒，Action 约 9.403 秒；编码、写入和解码各阶段已保留在 trace 中。

## 语义与保护验证

对最终实现与修复前函数在相同工作簿、相同 token salt 和工作表身份下直接比较完整快照 JSON（含 token）：单格、公式错误格、192 格混合内容区域、完整合并区域、合并区内部子区域均一致。

混合内容包含普通负数与错误值、布尔值、以等号开头的纯文本、公式、混合数字格式/字体/颜色/换行/对齐、隐藏行列及不同尺寸。

另验证：

- 完整合并区及周边普通格的九类请求格式逐格达标，原文字保留。
- 已失效 token 拒绝修改。
- 只覆盖部分合并区域的两种边界情况拒绝。
- 合并区内容写入仍拒绝。
- 受保护工作表及数组公式区格式化仍拒绝。

## 覆盖缺口与本轮限制

这两项属于执行侧应覆盖的错误：允许输入大小与默认时限不匹配，以及 Action 组合后的合法能力被通用检查拦截。此前“小范围单 Action 成功”不能证明近上限读取及“先合并再格式化”组合正确；不能把它们归咎于 Agent。

最终统计只包含最终验证版本的七个任务。早期优化版本仍超时、测试夹具调用错误和副本同名占用均保留在各自诊断目录，不混入最终成功率。本轮没有重新执行用户提示词到 Agent 编排的端到端测试，没有宣称所有未来文档都能在固定时间内完成。

## 交付与复跑入口

修改源码：

- `src/main/resources/wps_skills/excel/windows/excel_actions.ps1`
- `src/main/python/wps_skills/excel/contracts.py`
- 由契约生成的 Excel Action 定义及 `references/schemas/actions/formatRange.json`

新增可复跑的 Windows 执行回归：

- `src/test/python/tests/applications/range_acceptance.py`：直接 Task JSON，保存完整回执、耗时 trace；输入文件命名规则见模块说明。
- `src/test/resources/wps_skills/excel/range_regression.ps1`：在自建工作簿验证快照及合并保护；可用 `-LegacyActions` 指定旧函数进行语义对照。

本地证据：`build/archive/runs/excel-range-fix-20260919-final/evidence/`，含最终 `report.json`、trace、文档产物与 `installed.json`；语义对照报告位于上一级 `semantic-report.json`。

Windows 已安装位置：`C:/Users/yim/.workbuddy/skills/wps-excel`。只替换已验证的 `excel_actions.ps1`、`contracts.py` 和 `formatRange.json` 三个文件，安装前后哈希已核对。旧文件备份位于 `C:/Users/yim/wps-excel-range-fix-20260919-final2/installed-backup`。
