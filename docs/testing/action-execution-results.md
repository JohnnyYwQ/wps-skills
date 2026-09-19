# 全 Action 执行端验收结果

2026-09-19，在 Windows 已登录桌面完成固定 Task Request JSON 验收。最终独立轮次 **v6 全部符合预期，覆盖 85/85 个正式 Action**。这是执行端基线，不是 Agent 端到端测试，也不是所有参数组合无缺陷的证明。

已有计时 trace 的离线复算见 [正常成功任务执行耗时](successful-task-performance.md)：排除两条素材准备任务后，27 个业务 Task、639 次 Action；Word/Excel/PPT Task 中位数分别 2.16/8.92/3.73 秒。仅作成功任务的执行速度描述，不更改下列完整验收计数与结果。

## 最终结果

| 应用 | Action 名称覆盖 | 正常 Task 成功 | 原生预期失败 Task | 执行前拒绝请求 | 实际 Action 结果（成功 / 预期失败） |
| --- | ---: | ---: | ---: | ---: | ---: |
| Word | 14/14 | 3/3 | 0 | 14/14 | 21 / 0 |
| Excel | 34/34 | 13/13 | 1/1 | 34/34 | 265 / 1 |
| PPT | 37/37 | 13/13 | 1/1 | 37/37 | 409 / 1 |
| 合计 | **85/85** | **29/29** | **2/2** | **85/85** | **695 / 2** |

共 116 次请求提交：85 次未受理的参数拒绝、31 个受理执行的 Task。每个正式 Action 至少一次成功执行并通过明确的 data 字段断言。未执行步骤不计作失败 Action；预期失败也不计作正常成功。

- Excel X12 在 `stale_write` 返回 `STALE_RANGE`；PPT P12 在 `stale_edit` 返回 `STALE_CONTENT`。两者的后续步骤与持久化均停止，未产生计划末尾的输出文件。
- 85 个拒绝用例分别给一个 Action 注入未声明参数，验证完整计划提前拒绝、没有受理身份、没有执行步骤、输入文件保留。这不等于已经穷尽类型、范围和语义错误。
- 已受理 Task 的 Action/Task Response、退出码、输入消费、持久回执查询和 cleanup 均符合预期。29 个成功 Task 的已保存测试文档按精确路径关闭；两份预期失败现场保留。没有退出 WPS 或关闭其他文档。

## 测试条件与可复核证据

使用当前未提交工作区构建的独立测试包，Git 基准 `9ec6821da6df136bc8faf6fc89974f9c591fdb41`；工作区内容以包内逐文件哈希为准，不能只用这个 commit 还原。

- Windows 版本 `10.0.26200.0`，WPS `12.1.0.28505`，Python `3.13.14`，Windows PowerShell `5.1.26100.9444`、64 位。
- Mac 经 SSH 上传后，计划任务在 Windows 交互桌面 Session 3、普通权限启动 Python；Python 提交真实独立 Skill CLI，通过正式 bridge 和 WPS COM 执行。
- 测试目录为远端用户目录下 `wps-action-acceptance-v6/runs/{word,excel,ppt}`。没有替换宿主正在使用的 Skill。
- 运行窗口为北京时间 19:19:58–19:24:05。单应用用例累计约 Word 9.26 秒、Excel 170.80 秒、PPT 65.40 秒，包含 CLI、验证、回执和测试文档关闭；机器已有 WPS 进程。这不是冷启动性能或纯通信耗时，不能用作方案优劣结论。

测试包：[wps-action-acceptance.zip](../../build/distributions/wps-action-acceptance.zip)。SHA-256：`747f2bbf9526137c6d0dffb16cf4eaa2eccc1e9f498df89eba1701a1988154b8`。解压即有 `run_word.py`、`run_excel.py`、`run_ppt.py` 和三份独立待测 Skill。详细操作见 [执行端验收方案](action-execution-plan.md)。

本地证据：

- [汇总及传输哈希核验](../../build/runs/action-acceptance/v6/verified-summary.json)
- [Word 报告](../../build/runs/action-acceptance/v6/word/report.json)、[Excel 报告](../../build/runs/action-acceptance/v6/excel/report.json)、[PPT 报告](../../build/runs/action-acceptance/v6/ppt/report.json)
- 每个应用目录都有 `manifest.json`、`requests/`、`expected/`、`results/`、`receipts/`、`traces/` 和 `outputs/`。每个用例的响应和断言差异单独保存，可不重跑 WPS 就进行检查。
- 远端逐文件 SHA-256 清单与三份原始 ZIP 同存于 v6 目录；本地 1,319 个文件与远端清单完全匹配。

## 实际检查深度

Task 级检查状态、停止位置、身份、清理、持久回执与输入消费；Action 级检查顺序、响应身份、错误码、数据精确值或容差，并按正式结果契约核验。具体业务场景和每 Action 覆盖矩阵见方案及各应用 `manifest.json`。

文件检查覆盖存在性、大小、哈希、OOXML CRC 与主 XML、PDF/PNG 签名。Word 还独立检查替换/追加文字、表格内容和嵌图。Excel/PPT 多数业务效果目前依赖 Action Response 逐字段断言，文件容器检查不能代替独立业务内容验收。**人工视觉验收仍为 pending。**

本地回归通过 236 项测试，其中 3 项平台条件跳过，见 [测试日志](../../build/evidence/action-acceptance-tests.log)。取回证据时另修复 Windows ZIP 反斜杠目录项在 Mac 解压的问题，先复现失败，再通过 16 项诊断工具测试；此修复没有改变 v6 测试包或实机结果。

## 早期失败与修复

不能把最终通过覆盖到先前轮次。v2–v5 均在首个非预期失败时停止，没有重放原 Task：

| 轮次 | 停止位置 | 根因与修复 |
| --- | --- | --- |
| v2 | Word W01 | 测试预期多写了末尾段落标记；改为契约中的精确逻辑文本。 |
| v3 | Word W01 | 保留的前轮文档与新轮输出同名，WPS 正确拒绝；文件名加入完整运行目录哈希。 |
| v4 | Word W01 | 测试关闭脚本的 PowerShell 表达式枚举了 COM 集合；修复集合赋值。 |
| v5 | Word W02 | 测试把打开时的历史文件大小与随后保存后的大小比较；仅按保存/导出的最终产物检查。 |

详见 [失败台账](../../build/runs/action-acceptance/harness-failures.json)。这些是测试夹具、断言或测试清理问题；此次没有为使验收通过而改动生产 Action 实现。v6 使用修正后的同一份包重新完整执行三应用。

## 下一阶段

可以在本轮已验证范围内开始 [设计选择与价值验证](design-value-plan.md)，将 Agent 选择填参、宿主环境、Task 机制和 Action 失败分别归因。后续补齐参数分支与独立文件内容检查；响应丢失、超时隔离、焦点切换等故障实验另设判断标准，不混入本轮正常执行成功率。
