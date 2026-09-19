# Excel X07 多行插入失败：根因与修复

2026-09-18。根因已通过 Windows 交互桌面对照实验确认：原生脚本使用 `if` 的输出为变量赋值，PowerShell 将可枚举的多行 COM Range 展开为 `System.Object[]`，随后调用 `.Insert()` 的对象类型错误。不是 WPS 缺少多行插入能力，也不是 JSON、公式或 Task 编排错误。

## 复现与缩小

原正式轮次 X07 在 `insert_rows`（start=3、count=2）报 `RANGE_WRITE_FAILED`，消息为“找不到‘Insert’的重载，参数计数为：‘0’”。Action/Task 为 unknown，后续停止。

通过旧安装包，将请求缩为新建空白表、列出工作表、获取工作表凭据、insertRows，无数据、公式、格式、保存。交互桌面两例结果：

| 插入数量 | Task outcome | 耗时 | 结果 |
| --- | --- | ---: | --- |
| 1 行 | succeeded | 2.098 秒 | 成功 |
| 2 行 | unknown | 1.503 秒 | 复现相同错误 |

[最小复现原始响应](../../../build/archive/runs/excel-x07-diagnosis/minimal-report.json)。复现通过包内正式 Excel CLI 执行；既有失败文档没有被重放。

## 三个假设与对照证据

按顺序验证：① if 输出枚举导致类型变化；② 多行 COM Insert 要求显式可选参数；③ 范围地址构造错误。

在每例独立新建的临时工作簿中，将 A3 写为 sentinel，获取 A3:A4 的 EntireRow；保持相同零参数 Insert，只改变赋值方式：

| 赋值方式 | 实际类型 | 范围 | 插入后 A5 |
| --- | --- | --- | --- |
| `$axis = if (...) { $range.EntireRow } ...` | System.Object[]，长度 2 | 数组第一项为第 3 行 | 调用失败 |
| `if (...) { $axis = $range.EntireRow } ...` | System.__ComObject | 第 3～4 行 | sentinel，成功 |
| if 分支输出前加一元逗号 | System.__ComObject | 第 3～4 行 | sentinel，成功 |

[实验脚本](../../../build/archive/runs/excel-x07-diagnosis/probe.ps1)与[原始类型及效果记录](../../../build/archive/runs/excel-x07-diagnosis/probe-result.json)。实验只关闭自己刚新建的临时工作簿，没有关闭用户文档。

这确认第一个假设；正确 COM 对象上零参数 Insert 正常，排除第二个假设；范围地址和行数正确，排除第三个假设。一行的输出被折叠为单个 COM 对象，两行变数组，解释了旧单行回归通过而复杂场景失败。

## 修复范围

`src/main/resources/wps_skills/excel/windows/excel_common_actions.ps1`：将赋值放进 if/else 分支，保留整个 COM Range 对象。四种行列插删共用这一段，因此同时覆盖 insertRows、deleteRows、insertColumns、deleteColumns。

不修改 Action 参数、schema、验证条件或 Task 重试规则。原本在原生调用前已将操作标为“可能产生效果”，故该异常被保守报告为 unknown；本次没有为了提高成功率将 unknown 改为 failed。

独立构建包 [wps-excel](../../../build/archive/runs/excel-x07-fixed/wps-excel) 与前次实测包仅两处文件差异：上述脚本及 runtime/files.sha256.json。未替换或覆盖原始复杂测试证据，未发布 GitHub Release。

新增原生回归用例生成器 [structure_acceptance.py](../../../src/test/python/tests/applications/structure_acceptance.py)，覆盖 1/2/3 行、1/2/3 列插入后删除；检查插入空白位置、存量值位移、删除后的原矩阵与公式恢复。另以新路径重跑原 X07 完整请求。回归结果见下方补充。

## 修复后实机验证

Windows 交互桌面、独立文件路径，7/7 Task、121/121 Action 成功；全部关键响应断言通过，7/7 回执查询一致，输入文件消费成功，cleanup 全部成功。

| 用例 | 结果 | CLI 墙钟秒 |
| --- | --- | ---: |
| Rows1 | Task / Action / 返回值断言通过 | 17.258 |
| Rows2 | Task / Action / 返回值断言通过 | 17.340 |
| Rows3 | Task / Action / 返回值断言通过 | 18.091 |
| Columns1 | Task / Action / 返回值断言通过 | 16.186 |
| Columns2 | Task / Action / 返回值断言通过 | 17.161 |
| Columns3 | Task / Action / 返回值断言通过 | 17.970 |
| X07 | Task / Action / 返回值断言通过 | 34.907 |

原 X07 的全部 25 个 Action 成功，插删后数据恢复原矩阵、合计公式恢复 `=SUM(B2:B7)`、值为 210，保存成功。它是修复后独立验证，不改写原来正式轮次的 21/22 正常 Task 成功率。

本地 Task 与复杂请求测试 11/11 通过。[完整修复证据 ZIP](../../../build/archive/loose-files/excel-x07-fixed-evidence.zip)、[全部回执](../../../build/archive/runs/excel-x07-fixed-evidence/report.json)、[逐例断言结果](../../../build/archive/runs/excel-x07-fixed-evidence/comparison.json)。最小实验和对照脚本保存在 build/archive/runs/excel-x07-diagnosis，作为诊断证据；生产代码未加入临时调试日志。
