# 三应用复杂业务执行测试 R07

2026-09-19，在同一 Windows 交互桌面依次运行 `run_word.py`、`run_excel.py`、`run_ppt.py`。本轮按固定请求和预期一次完成，中途没有修改脚本或重跑用例。

## 完成情况

60 个业务 Task 全部执行，14 个素材准备 Task 全部通过，无 blocked、not_run 或仍在运行的业务用例。启动器从 UTC 14:51:51.162 到 14:59:42.597，共 **471.43 秒（7 分 51 秒）**；该时间包含准备、CLI 启动、回执查询及测试文档收尾，不等于业务执行时间之和。

| 应用 | 原始业务断言通过 | Task Response | Action Response | 清理成功 |
| --- | ---: | --- | --- | ---: |
| Word | **19/20** | 20 succeeded | 330 succeeded | 20/20 |
| Excel | **18/20** | 18 succeeded、2 failed | 407 succeeded、2 failed；11 个后续步骤未执行 | 20/20 |
| PPT | **18/20** | 18 succeeded、2 unknown | 705 succeeded、2 unknown；36 个后续步骤未执行 | 20/20 |
| 合计 | **55/60（91.67%）** | 56 succeeded、2 failed、2 unknown | 1442 succeeded、2 failed、2 unknown；47 个步骤未执行 | **60/60** |

这些是固定 JSON 的执行侧数据，不是 Agent 自主任务完成率。85 个正式 Action 名称均在本轮通过用例中至少有一次成功验证，但这不能证明全部参数组合都正确：PPT 的新组合仍暴露了回读不一致。

## 五个未通过用例的分析

| 用例 | 原始现象 | 归因与处理结论 |
| --- | --- | --- |
| CW15 风险章节扩充 | Task 及全部 Action 成功，`risks_read.text` 比预期多一个末尾换行 | 测试断言错误。插入完整段落的返回范围含末尾分段符，正文、字体及后续保存均通过。仅对已保存响应补充该精确边界值做离线诊断后，全部响应断言通过；原始 19/20 不改写。 |
| CE13 库存盘点 | `sort` 返回 `RANGE_UNSUPPORTED` | 测试请求超出当前能力。数量计算后直接对含公式的区域排序，而 `sortRange` 的正式用途限定为常量值行。应先生成数值快照，再排序快照；不能要求当前执行端对这个请求成功。 |
| CE15 培训成绩 | `rank` 返回同一 `RANGE_UNSUPPORTED` | 同一请求编排问题：平均分、合格标记仍是公式，随后直接排序。应将排名区固化为数值后排序，保留计算表。 |
| CP16 路线图 | `stage_style0` 返回 unknown / `INVALID_RESULT` | 正式契约接受的文本框样式参数在实机回读中不满足要求，需跟进执行端兼容性。 |
| CP17 组织职责 | `style0` 返回同一 unknown / `INVALID_RESULT` | 与 CP16 相同触发组合，不能认定为两个独立根因。 |

PPT 两例共同条件：新建文本框，修改前 `lineVisible=false`，观察到 `lineWidth=-2147483904`；提交的 patch 为 `fillColor=15921906, lineWidth=1`，未显式改变边框可见性。原生操作返回后，正式结果语义校验判定“shape style readback differs from requested patch”。当前失败响应没有保留实际修改后的 appearance，**只能定位到该参数组合，不能仅凭这份记录断言究竟哪个属性未生效**。

同批 CP14 在隐藏边框文本框上仅设置填充色并保持边框隐藏通过，因此不是所有文本框填充都失败。下一步可单独对比隐藏/显示边框时的线宽写入和回读。本轮没有为消除失败而临时改参数、移除断言、回放 Task 或修改产品代码。

三个测试用例问题属于我的测试设计/断言问题；PPT 两例保留为执行端待定位项。当前不能得出“执行端全部没有问题”的结论。

## 正常成功业务耗时

严格采用原始断言通过的 **55 个业务 Task**。不纳入 setup、原始失败用例或未执行步骤，不删除慢但成功的任务。Task 为原 trace 的 `request.total`，Action 为 `action.execute`，后者已包含在前者中。

| 应用 | 样本 | Task 中位数 | Task 范围 | Task 累计 | 成功样本内 Action 次数 | Action 中位数 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Word | 19 | **2.54 s** | 1.82–12.69 s | 60.66 s | 319 | 100.91 ms |
| Excel | 18 | **9.61 s** | 2.26–48.23 s | 227.21 s | 385 | 201.73 ms |
| PPT | 18 | **3.91 s** | 1.94–6.72 s | 74.49 s | 693 | 47.44 ms |

这些是不同工作量的固定业务场景各一次执行。时间包含实际读写、内部验证和计划中的观察操作，不含 Agent 推理；不能作为设计对照提速比例或稳定性结论。

## 三个脚本的实际行为

本轮已验证用户要求的自动运行方式：Word 的失败未阻塞余下 Word 用例或 Excel；Excel 的两个失败后继续完成其余用例并进入 PPT；PPT 的两个 unknown 在正式清理成功后，后续独立任务正常通过。失败 Task 内部的后续 Action 仍停止；测试工具继续的是另一份独立文档的 Task。

全部业务 Task 的 cleanup 为 succeeded。Windows 专用计划任务结束后为 Ready，未继续后台执行。失败文档保留，不强制丢弃；已有轮次记录不覆盖。

## 版本与证据

- 测试包：`build/tools/complex-business-20260919-r07.zip`，包含且仅包含三个顶层 Python 入口及公共库。
- 测试包 SHA-256：`aefdad4a611af9d15bff625fe7af834a56696d399f8100f555e9eb7f23e839fd`。
- Windows 原始证据 ZIP SHA-256：`e77d38c974cfe45622f2552787c8bb6eba339573667d127745ceee975eabf765`；下载后核对 **1407 个文件**的远端 SHA-256 全部一致。
- 本次前后 `src/main` 的 254 个生产文件哈希一致；没有将测试脚本修正计为产品修复。
- [逐 Task / Action 与应用汇总](../../build/evidence/complex-business-20260919-r07/summary.json)
- [逐 Task CSV](../../build/evidence/complex-business-20260919-r07/tasks.csv)
- [逐 Action CSV](../../build/evidence/complex-business-20260919-r07/actions.csv)
- [CW15 离线边界诊断](../../build/evidence/complex-business-20260919-r07/cw15-boundary-analysis.json)
- [原始证据目录](../../build/runs/complex-business-20260919-r07)
- [脚本复审记录](complex-task-script-review.md)

本轮到此结束，不启动新轮次。原始断言通过率与诊断结论分别保留。
