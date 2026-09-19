# Word Skill 进阶实机测试结果

日期：2026-09-17。Windows 交互桌面 Session 1，使用 `build/archive/runs/word-fingerprint-fix-20260917-223921/wps-word`；56 个包文件清单校验通过。

**8 项全部执行；6 项满足清单全部预期，2 项响应值存在差异。** 8 个 Task 和 97 个 Action 均返回 succeeded，无 Action error、无 Task stop；每项仅提交一次，没有重试。人工粗审待进行。

| 编号 | 用例 | 响应断言 | 差异所在 Action | 人工粗审 |
| --- | --- | --- | --- | --- |
| 19 | 三节布局与连续 revision 引用 | 通过 | 无 | 未审阅 |
| 20 | 三节六变体页眉页脚与继承 | 响应值与清单预期不符 | base / inspectDocument：pageBreakCount 预期 6，实际 8 | 未审阅 |
| 21 | 重复查找、变长替换与相对插入 | 通过 | 无 | 未审阅 |
| 22 | 混合格式段落替换与删除 | 通过 | 无 | 未审阅 |
| 23 | 多张表格与正文连续插入 | 通过 | 无 | 未审阅 |
| 24 | 多图片尺寸模式与分页混排 | 响应值与清单预期不符 | image2 / insertImage：预期 100×50 pt（±0.1），实际 99.75×49.5 pt | 未审阅 |
| 25 | 编辑独立复杂文档并另存 | 通过 | 无 | 未审阅 |
| 26 | 三节综合报告与较长 Task | 通过 | 无 | 未审阅 |

## 20：分页统计差异

- 步骤：`base`，Action：`inspectDocument`。
- 清单预期：`data.structure.pageBreakCount=6`；实际：`8`。`sectionCount=3`、`sectionBreakCount=2` 均符合预期。
- 请求确实插入 6 个 page 和 2 个 sectionNextPage。当前 [word_bridge.ps1:1283](../../../src/main/resources/wps_skills/word/windows/word_bridge.ps1#L1283) 通过统计正文中的 `\f` 计算 pageBreakCount，未在该计算中区分显式分页与分节。数值与“6+2”吻合，提示统计口径混入分节标记；此次未额外进行原生 COM 定位，不能把此线索写成已完成根因复现。
- 后续六变体页眉页脚、继承、独立覆盖、清空、保存和 PDF 均成功且满足其余响应断言。这不是某个 Action 抛错或内容未执行；不代表实际多出两页。

## 24：contain 图片尺寸差异

- 步骤：`image2`，Action：`insertImage`。
- 输入：200×100 PNG，`size.kind=box`，100×100 pt 方框，`fit=contain`。
- 清单预期：100×50 pt，容差 0.1 pt；实际：99.75×49.5 pt，宽差 0.25 pt、高差 0.5 pt。图片源、内嵌方式、替代文字均匹配；另外两张图片及保存/PDF 通过。
- 当前 [原生尺寸回读](../../../src/main/resources/wps_skills/word/windows/word_actions.ps1#L1577) 和 [结果验证](../../../src/main/python/wps_skills/word/contracts/validation.py#L414) 对 contain 仅要求不超过目标框（上界容差 0.5 pt），没有要求恰好达到按源宽高比推算的 100×50。因此此次成功响应符合当前校验口径，但未满足测试单较严格的尺寸预期。
- 尚不能据此认定缩放实现错误，或直接认定只是 WPS 量化误差；需进一步确认尺寸承诺与实际精度。本轮保留差异，不修改断言、不修改执行代码。

## 第 25 例独立素材

准备 Task ID：`task-089420c97168166064e8bce15699b85e26450590c9f899cc7c9c1ddb89e813ef`，独立创建文本、2×2 表格、图片、两节及页眉页脚并保存，再复制为该例自己的输入路径。准备请求和响应分别保存在 setup.original.json / setup.stdout.json；准备不计入正式 8 项通过数。

25 的 openDocument 和 initial 均为 saved，初始两节纵向、表格和图片计数、primary 页眉页脚文字与继承、变体关闭状态均符合素材约定。编辑后 modified，另存和 PDF 成功。

## 人工审核与文件

自然语言效果仍以 [进阶测试清单](WPS_WORD_ADVANCED_TEST_CHECKLIST.md) 各例为准；两项差异没有被改写成新的效果预期。按照约定没有额外解析最终 DOCX/PDF、截图或视觉验收。

Windows 根目录：`C:/Users/yim/wps-advanced-20260917-231026`。outputs 下产物名为 `20260917-231026-output-编号.docx`；20/24/25/26 同时有 PDF。测试文档保留，临时计划任务已移除。

- [原始汇总报告](../../../build/evidence/word-advanced-20260917-231026/windows/report.json)
- [完整 Action/Task 响应](../../../build/evidence/word-advanced-20260917-231026/windows/responses)
- [原始请求](../../../build/evidence/word-advanced-20260917-231026/windows/original-requests)
- [产物](../../../build/evidence/word-advanced-20260917-231026/windows/outputs)
- [响应核对脚本](../../../build/evidence/word-advanced-20260917-231026/windows/harness/verify.py)
- [运行前清单存档](../../../build/evidence/word-advanced-20260917-231026/checklist-before-run.md)

本轮说明正常复杂组合大部分可执行，但已出现两处测试预期与实际响应的差异，不能宣布进阶 8 例全通过，更不能据此把后续所有问题归因于 Agent。


最新统一复测：三阶段38例已使用正式trace包重新执行，结果和Task/Action/通信耗时见 [三阶段统一报告](WPS_WORD_THREE_STAGE_RESULTS.md)。本文件上文保留为历史轮次记录。
