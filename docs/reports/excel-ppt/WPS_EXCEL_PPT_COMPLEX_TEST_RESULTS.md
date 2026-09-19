# Excel / PPT 复杂场景实机测试结果

2026-09-18。正式交互桌面轮次 R02 已完成：**正常 Task 21/22 成功，正常场景已执行 Action 583/584 成功；两项保护用例均通过，全部用例 23/24 符合预期。** 尚未人工粗审，不是 Agent 端到端测试。

## 环境与口径

Windows 10.0.26200，WPS 12.1.0.28505，Python 实际版本 3.13.14，交互桌面 Session 2。测试根目录为 `C:/Users/yim/wps-complex-plan-20260918-r02`。安装目录名字含 3.13.12，不能据此推断解释器版本。

沿用此前迁移验证的 Excel/PPT 安装包，193 个原始包文件的 SHA-256 全部一致；不比较运行时产生的 pyc。24 份正式固定 JSON 各提交一次，另有 2 份输入准备 Task。未修改生产实现、未放宽断言容差、未自动重试失败 Action。X07 停止后确认其资源释放及进程退出，再继续剩余独立文档用例。

最开始 R01 错从 SSH 会话启动，未先验证交互桌面，是本次测试入口疏漏。该轮使用 Python 3.10.18，X10 返回只读、P10 打开超时；其证据单独保留，不并入正式成功率或耗时。修正入口后在新文件路径重测，R02 的 X10/P10 均成功，因此前两项现象目前属于环境相关异常，不能直接认定为应用执行缺陷。X07 在两轮中均复现。R01 非交互 WPS 进程仍存在，不曾强制关闭或清除隔离状态；本报告不是空机性能基准。

另修正一项预期文件错误：`taskFile.state` 应为 `removed`，原清单误写 `consumed`。依据正式代码修正生成器，原预期保留；R02 提交前已使用正确预期。这不是运行时缺陷。

## 成功率与返回值比对

| 指标 | Excel | PPT | 合计 |
| --- | ---: | ---: | ---: |
| 正常 Task 成功 | 10/11，90.91% | 11/11，100% | 21/22，95.45% |
| 正常场景已执行 Action 成功 | 224/225，99.56% | 359/359，100% | 583/584，99.83% |
| 正常场景全部响应预期达标 | 10/11 | 11/11 | 21/22 |
| 保护用例按预期拒绝并停止 | 1/1 | 1/1 | 2/2 |
| 全部用例符合预期 | 11/12 | 12/12 | 23/24，95.83% |

正式请求共计划 639 个 Action：612 succeeded、2 failed（均为故意触发的保护）、1 unknown（X07）、24 not_executed。正常场景有 15 步因 X07 停止而未执行，保护场景有 9 步按预期未执行，均不计入“已执行 Action”分母。

两份 setup 的 27 个 Action 全成功，单独统计。含 setup 的全部 26 份回执查询均与提交结果一致，输入均被消费，cleanup 均成功。全部成功场景没有“Action 成功但关键返回值断言不符”的情况。X07 在断言文件中引出 106 条连锁差异，来自一个停止点及其后续缺失结果，**不是 106 个独立缺陷**。

## 唯一正式未完成用例：X07

- 场景：已有 6 行明细与合计公式，插入两行、删除两行、插删一列，期望恢复数据及公式。
- 出错步骤：`insert_rows`，Action `insertRows`，参数 `sheet="结构调整"`、`start=3`、`count=2`（以归档 JSON 为准）。
- Task 与该 Action outcome 均为 `unknown`，错误码 `RANGE_WRITE_FAILED`，消息：`找不到“Insert”的重载，参数计数为:“0”。`
- 出错前 9 个 Action 成功（含新建文档），其后 15 个未执行，保存未执行；不能宣称两行一定未插入，也不能重放。
- cleanup 成功、拥有的桥接进程已退出；保留首次失败文档及回执，没有用修复重测覆盖它。
- 已确认调用失败位置是原生 `excel_common_actions.ps1` 的 `$axis.Insert()`。此前简单回归插入一行通过，本轮一次插入两行失败。多行范围的获取/展开与 COM 调用方式是下一步定位对象，底层成因尚未用最小实验确证；本次没有修复。

这说明此前“所有 Action 至少成功执行过”不能替代复杂输入覆盖。本轮扩充测试确实找到了现有简单场景未覆盖的问题。

## 逐例结果与时间

以下为 Windows 本机 CLI 进程墙钟秒数，包含启动 Python 到收齐响应，不含 SSH 部署、回执查询和人工审核。保护场景的 failed 是正确结果。

| 用例 | 场景 | Task outcome | 响应预期 | 总耗时（秒） |
| --- | --- | --- | --- | ---: |
| X01 | 三部门月度报表与汇总 | succeeded | 通过 | 78.132 |
| X02 | 公式链、条件判断与受控错误 | succeeded | 通过 | 25.846 |
| X03 | 混合类型与分区格式 | succeeded | 通过 | 19.395 |
| X04 | 排序、筛选与恢复显示 | succeeded | 通过 | 31.734 |
| X05 | 跨表复制数值、公式值固化与源数据保留 | succeeded | 通过 | 34.463 |
| X06 | 工作表复制、重命名、移动与删除 | succeeded | 通过 | 15.442 |
| X07 | 行列插删后的数据与公式调整 | unknown | 未通过 | 9.547 |
| X08 | 分组标题合并、取消合并与重排 | succeeded | 通过 | 21.170 |
| X09 | 查找替换与公式保护 | succeeded | 通过 | 9.820 |
| X10 | 已有文件编辑与原路径保存 | succeeded | 通过 | 11.990 |
| X11 | 较大明细、公式列与双格式交付 | succeeded | 通过 | 154.727 |
| X12 | 复杂链中的过期凭据与停止 | failed | 通过 | 8.928 |
| P01 | 六页项目启动演示 | succeeded | 通过 | 5.302 |
| P02 | 中西文字体、字号与连续改写 | succeeded | 通过 | 2.747 |
| P03 | 多段文字与文本框布局 | succeeded | 通过 | 3.145 |
| P04 | 多形状对齐与等间距分布 | succeeded | 通过 | 2.924 |
| P05 | 层叠形状、外观与几何调整 | succeeded | 通过 | 2.637 |
| P06 | 幻灯片复制、移动、删除与身份保持 | succeeded | 通过 | 2.574 |
| P07 | 三页图文表混排 | succeeded | 通过 | 3.614 |
| P08 | 双表格与连续矩阵更新 | succeeded | 通过 | 5.941 |
| P09 | 中文查找替换与讲者备注 | succeeded | 通过 | 2.360 |
| P10 | 已有三页文稿的局部修订 | succeeded | 通过 | 2.035 |
| P11 | 五页演示与 PPTX/PDF/PNG 交付 | succeeded | 通过 | 5.308 |
| P12 | 复杂文稿中途失败后的停止 | failed | 通过 | 2.033 |

## 耗时分析

成功正常 Task：Excel 10 例中位数 **23.51 秒**，最长 **154.73 秒**；PPT 11 例中位数 **2.92 秒**，最长 **5.94 秒**。每例只测一次，不是稳定延迟或 SLA；样本小，最近秩 P95 在这里等于最大值。不能把不同复杂度任务混成“每个任务平均只需几秒”。

以下累计覆盖每个应用的 12 份正式请求（包含保护用例及 X07），不含 setup。各行存在包含关系，不能相加：

| 阶段 | Excel 累计秒 | PPT 累计秒 |
| --- | ---: | ---: |
| CLI 进程墙钟 | 421.194 | 40.622 |
| JSON 输入至响应写出 request.total | 420.450 | 39.876 |
| Action 执行 action.execute | 417.712 | 37.004 |
| bridge 往返 | 417.432 | 36.638 |
| bridge 编码 + 写入 + 解码 | 0.277 | 0.422 |
| bridge 等待返回 | 416.655 | 35.430 |
| Action 关联的 native.operation | 409.018 | 25.443 |
| Task cleanup | 0.645 | 0.610 |

Excel 的 Action 执行占内部请求总耗时约 99.35%，PPT 约 92.80%。Excel 时间主要落在原生操作，包括内容读取、写入及内置回读验证；Python 的编码、写入、解码不是当前主要成本。

**等待返回和 bridge 往返不能称作纯通信耗时。** 按 `(bridgeRequestId, operation)` 配对原生阶段后，Excel 504 次、PPT 798 次往返均能匹配。往返扣除已记录原生阶段后的未归因差值分别约 5.63 秒、7.22 秒；其中包含调度、传输、日志和未覆盖区间，只能称为桥接及未归因开销估计。建立文档的一次 Action 会有多个原生 operation，不能只按 bridgeRequestId 合并再重复相减。

最慢的主要步骤都在 X11（40 条、8 列）：

| Action / 步骤 | 秒 |
| --- | ---: |
| writeRange / bulk_data | 46.226 |
| saveAs / task_save | 25.910 |
| exportPdf / task_pdf | 25.708 |
| readRange / bulk_data_token | 24.090 |
| setFormulas / bulk_formula | 9.491 |

X11 全 Task 154.73 秒，单个 Action 未超过 60 秒上限。这提示后续应测量并优化逐单元格观察、写入和保存/导出的观察成本；现有数据还不能细分原生阶段内部每个 COM 调用的占比。

## 人工粗审与证据

[人工粗审清单及文件链接](../../../build/archive/runs/excel-ppt-complex-r02-evidence/HUMAN_REVIEW.md)逐例提供自然语言预期和可打开文件；全部标为待审核。Windows 产物目录为 `C:/Users/yim/wps-complex-plan-20260918-r02/outputs`，本地副本见 [outputs](../../../build/archive/runs/excel-ppt-complex-r02-evidence/outputs)。X07 和两个保护场景没有保存交付文件，不应据此判定输出丢失；其保存步骤应停止。

- [正式原始证据 ZIP](../../../build/archive/loose-files/excel-ppt-complex-r02-evidence.zip)：请求、预期、原始 stdout、持久回执、trace、产物。
- [正式结果摘要](../../../build/archive/runs/excel-ppt-complex-r02-evidence/summary.json)与[原始全部响应](../../../build/archive/runs/excel-ppt-complex-r02-evidence/report.json)。
- [逐 Task 耗时](../../../build/archive/runs/excel-ppt-complex-r02-evidence/tasks.csv)、[逐 Action 结果及耗时](../../../build/archive/runs/excel-ppt-complex-r02-evidence/actions.csv)、[所有计时跨度](../../../build/archive/runs/excel-ppt-complex-r02-evidence/timings.csv)、[原生阶段](../../../build/archive/runs/excel-ppt-complex-r02-evidence/native-timings.csv)。
- [响应断言差异](../../../build/archive/runs/excel-ppt-complex-r02-evidence/assertions.csv)、[按 Action 聚合的耗时](../../../build/archive/runs/excel-ppt-complex-r02-evidence/timing-summary.json)、[桥接配对统计](../../../build/archive/runs/excel-ppt-complex-r02-evidence/bridge-comparison.json)。
- [包一致性验证](../../../build/archive/runs/excel-ppt-complex-r02-evidence/package-verification.json)、[环境记录](../../../build/archive/runs/excel-ppt-complex-r02-evidence/environment-final.json)。
- [环境异常轮次 R01 原始证据](../../../build/archive/loose-files/excel-ppt-complex-r01-evidence.zip)，不并入正式指标。

本轮临时调度任务已删除。未关闭用户文档、未清除不确定文档的保护、未声称人工审核通过。下一步应先定位并修复 X07 多行插入，再以新的独立文件验证多行/多列插删及原复杂链；Excel 的大区域性能可另立基准，避免与功能正确性混为一个结论。

## 后续修复补充：X07

X07 根因已经确认并修复：PowerShell 的 if 输出将多行 COM Range 展开成 Object[]。改为分支内直接赋值后，原 X07 在新路径下 25/25 Action 成功；另有六组 1/2/3 行及列插删验证通过，共 7/7 Task、121/121 Action 与响应预期均通过。详见 [根因与修复报告](WPS_EXCEL_X07_DIAGNOSIS.md)。上方保留原始整轮指标；这不是修复后将全部 24 例又测一遍的结果。
