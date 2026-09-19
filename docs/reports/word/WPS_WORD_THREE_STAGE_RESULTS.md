# Word 三阶段统一实机测试与耗时分析

2026-09-17，Windows WPS 12.1.0.28505、交互桌面 Session 1。统一使用 `build/archive/runs/word-timing-final-20260917/wps-word`，58 个包清单文件校验通过。38 例按 01→38 顺序，各提交一次，使用全新目录与唯一文件名。所有字体按用户要求假定已安装。

**38/38 Task 成功，226/226 已执行 Action 成功；36/38 清单响应断言通过。** 无 Action failed/unknown、无未执行 Action；38 例 cleanup 均成功。第20、24例保持原有断言差异，没有把它们当作Action报错，也没有放宽断言改判通过。

素材准备另有一个Task，未计入38例和所有性能统计。第一阶段各已有文件为独立原始素材；25的复杂素材独立准备并保存再复制。没有 Agent、重试或并发测试，没有修改图片、分页计数或生产执行逻辑。

## 成功率

| 阶段 | 范围 | Task成功 | Action成功 | 清单断言通过 |
| --- | --- | --- | --- | --- |
| 1 | 基础01–18 | 18/18（100%） | 80/80（100%） | 18/18（100.00%） |
| 2 | 进阶19–26 | 8/8（100%） | 97/97（100%） | 6/8（75.00%） |
| 3 | 文字格式27–38 | 12/12（100%） | 49/49（100%） | 12/12（100.00%） |

全组清单断言通过率为94.74%。Action成功率分母为实际执行的Action，本轮恰好等于计划的226个，包含document和completion；不会把未执行项算成成功。

第三阶段12例覆盖多级标题、单段多字号/强调、同run及跨run中西字体、四种对齐、六种行距、字号与行距交叉、悬挂/首行/左右缩进、多次追加显式恢复格式、完整段落替换、跨run查找替换、综合文字报告。全部通过Action内部回读及响应断言；人工粗审尚未进行，不据此宣称所有视觉排版已经确认。

保留差异：

- 20：`base / inspectDocument` 的 pageBreakCount仍为8，预期6；sectionBreakCount=2。后续页眉页脚、保存、PDF成功。
- 24：`image2 / insertImage` 回读99.75×49.5 pt，清单要求100×50 pt、容差0.1 pt。符合当前contain不超目标框的校验口径，但不满足较严格的测试断言。其余步骤成功。

## Task执行速度

内部Task为CLI开始处理输入到Task Response写出/flush结束；外部端到端从启动提交子进程到返回/退出，不含断言与报告处理。两者不可混用。

| 阶段 | 内部总计 s | 内部平均 s/Task | 内部中位 s | 内部最慢 s | 外部平均 s/Task |
| --- | --- | --- | --- | --- | --- |
| 1 | 34.608 | 1.923 | 1.864 | 3.108 | 1.993 |
| 2 | 26.676 | 3.335 | 2.816 | 6.121 | 3.406 |
| 3 | 28.315 | 2.360 | 2.250 | 2.800 | 2.430 |

38例内部总耗时89.599秒，外部总耗时92.283秒。Action累计82.942秒，Task其他开销6.656秒（7.43%），包括参数/回执处理、清理和输出等。外部与内部累计差2.684秒，平均约70.64 ms/Task，包含解释器/模块加载等外部启动开销。

最慢的是20（三节六变体，25个业务步骤），内部6.121秒；其次26（三节综合报告）4.868秒。复杂Task步骤更多，总耗时更长，不能直接将不同Task平均数作为性能退化证据。

## Action执行速度

以下为成功Action的execute调用耗时，包含桥接与内部回读；不是纯COM操作时间。

| Action | 次数 | 平均 ms | 中位 ms | 最慢 ms | 合计 s |
| --- | --- | --- | --- | --- | --- |
| createDocument | 28 | 1415.15 | 1401.20 | 1731.76 | 39.624 |
| exportPdf | 6 | 187.60 | 181.15 | 256.57 | 1.126 |
| findContent | 10 | 72.23 | 69.99 | 109.77 | 0.722 |
| insertBreak | 17 | 90.46 | 87.68 | 120.87 | 1.538 |
| insertImage | 7 | 146.37 | 155.85 | 184.78 | 1.025 |
| insertTable | 5 | 153.07 | 168.68 | 177.25 | 0.765 |
| inspectDocument | 21 | 274.75 | 285.96 | 509.91 | 5.770 |
| openDocument | 10 | 1288.68 | 1259.11 | 1590.59 | 12.887 |
| replaceContent | 16 | 117.93 | 120.66 | 158.82 | 1.887 |
| save | 1 | 89.01 | 89.01 | 89.01 | 0.089 |
| saveAs | 33 | 191.66 | 191.84 | 248.14 | 6.325 |
| setHeaderFooter | 7 | 245.85 | 213.22 | 409.11 | 1.721 |
| setPageLayout | 6 | 143.50 | 139.30 | 192.08 | 0.861 |
| writeContent | 59 | 145.81 | 124.57 | 349.11 | 8.603 |

主要发现：

- 新建/打开共38次，累计52.511秒，占Action累计耗时约63.31%。每个Task都建立独立执行资源，获取文档和桥接准备是主要固定成本。当前数据不支持将瓶颈归因于JSON传输，也不意味着应改变Task隔离模型。
- 排除createDocument/openDocument，其余188次Action平均约161.87 ms。inspectDocument平均274.75 ms，比简单文字查找72.23 ms更重，与其读取段落、runs、结构及指纹的工作范围不同相符；具体内部热点仍需更细粒度剖析才能确定。
- 第三阶段writeContent共17次，平均215.02 ms；第一/二阶段分别138.75/104.90 ms。第三阶段通常包含更多文字、格式项或段落，工作量不同，不能称为回归。全组最慢writeContent是02的格式样张349.11 ms；第三阶段最长为27多级标题326.52 ms。

## 通信与PowerShell执行

226/226 Action均有计时。566/566 Python桥接往返均匹配到PowerShell的三个阶段，无缺失、无负剩余耗时。跨端按 **bridgeRequestId + operation** 关联，并核对trace上下文；同一个Action获取文档时会复用bridgeRequestId执行多个不同operation，不能仅按该ID合并。素材准备日志不参与本统计。

| 阶段 | 桥接次数 | 往返累计 s | PowerShell operation累计 s | 扣除三个原生阶段后的剩余 s |
| --- | --- | --- | --- | --- |
| 1 | 214 | 31.792 | 23.370 | 7.105 |
| 2 | 218 | 25.164 | 21.153 | 3.364 |
| 3 | 134 | 26.413 | 21.033 | 4.616 |

全部桥接往返累计83.369秒，其中PowerShell operation累计65.556秒；PowerShell三个阶段合计68.284秒，剩余15.085秒。剩余包含启动/脚本加载、管道、排队、Python编码解码、系统调度和日志写入，**不是纯通信耗时**。

Python直接测量的566次编码累计99.89 ms，写入/flush累计95.85 ms，响应解码累计125.10 ms，三者合计320.84 ms。wait_response累计82.445秒包含WPS执行，不能将它全部算成通信成本。

38次prepare_new_document/prepare_existing_document是每个Task首次通信，其剩余累计13.628秒，占全部剩余约90.34%；每次平均约358.63 ms。根据调用顺序，这是惰性启动后的首次请求，尚待完成的PowerShell启动、脚本加载等会进入等待区间。此解释是由时序作出的推断，当前日志不能再把这部分精确拆成纯启动与传输。

排除首次prepare后，其余528次剩余累计1.457秒，平均2.76 ms/次。说明正常后续往返的额外开销较小；不是网络RTT，也不包含全部COM耗时。桥接累计可略高于Action累计，因为清理释放通信在Action之外；嵌套时间不可直接相加。

## 每例结果与耗时

| 编号 | 用例 | Task/Action | 清单断言 | 内部Task ms | 外部端到端 ms | 完整响应 |
| --- | --- | --- | --- | --- | --- | --- |
| 01 | 新建文字，不保存 | 成功 | 通过 | 1328.84 | 1398.53 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/01.stdout.json) |
| 02 | 文字与段落格式 | 成功 | 通过 | 1731.04 | 1801.49 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/02.stdout.json) |
| 03 | 在指定段落前插入 | 成功 | 通过 | 1879.35 | 1949.28 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/03.stdout.json) |
| 04 | 读取及查找 | 成功 | 通过 | 1825.53 | 1895.29 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/04.stdout.json) |
| 05 | 批量替换文字 | 成功 | 通过 | 1613.02 | 1682.90 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/05.stdout.json) |
| 06 | 替换指定完整段落 | 成功 | 通过 | 1911.20 | 1980.47 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/06.stdout.json) |
| 07 | 删除指定完整段落 | 成功 | 通过 | 1879.16 | 1950.76 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/07.stdout.json) |
| 08 | 插入表格 | 成功 | 通过 | 1907.14 | 1979.36 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/08.stdout.json) |
| 09 | 插入图片 | 成功 | 通过 | 1834.89 | 1902.86 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/09.stdout.json) |
| 10 | 页面方向与页边距 | 成功 | 通过 | 2089.05 | 2160.70 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/10.stdout.json) |
| 11 | 页眉页脚 | 成功 | 通过 | 2217.36 | 2287.13 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/11.stdout.json) |
| 12 | 插入分页符 | 成功 | 通过 | 1849.16 | 1918.02 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/12.stdout.json) |
| 13 | 分节与第二节横向 | 成功 | 通过 | 2312.87 | 2385.19 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/13.stdout.json) |
| 14 | 编辑已有文件并保存 | 成功 | 通过 | 1830.02 | 1899.18 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/14.stdout.json) |
| 15 | 编辑已有文件后另存 | 成功 | 通过 | 1759.17 | 1829.85 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/15.stdout.json) |
| 16 | 编辑已有文件，不保存 | 成功 | 通过 | 1639.35 | 1709.12 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/16.stdout.json) |
| 17 | 修改后仅导出 PDF | 成功 | 通过 | 1892.66 | 1963.24 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/17.stdout.json) |
| 18 | 完整报告：DOCX 与 PDF | 成功 | 通过 | 3108.01 | 3178.42 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/18.stdout.json) |
| 19 | 三节布局与连续 revision 引用 | 成功 | 通过 | 3046.41 | 3117.65 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/19.stdout.json) |
| 20 | 三节六变体页眉页脚与继承 | 成功 | 保留差异 | 6120.66 | 6199.06 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/20.stdout.json) |
| 21 | 重复查找、变长替换与相对插入 | 成功 | 通过 | 2222.12 | 2294.06 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/21.stdout.json) |
| 22 | 混合格式段落替换与删除 | 成功 | 通过 | 2421.65 | 2492.31 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/22.stdout.json) |
| 23 | 多张表格与正文连续插入 | 成功 | 通过 | 2252.99 | 2320.83 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/23.stdout.json) |
| 24 | 多图片尺寸模式与分页混排 | 成功 | 保留差异 | 2585.75 | 2653.76 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/24.stdout.json) |
| 25 | 编辑独立复杂文档并另存 | 成功 | 通过 | 3158.49 | 3229.06 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/25.stdout.json) |
| 26 | 三节综合报告与较长 Task | 成功 | 通过 | 4868.05 | 4939.59 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/26.stdout.json) |
| 27 | 多级标题与正文交错 | 成功 | 通过 | 2197.65 | 2269.46 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/27.stdout.json) |
| 28 | 单段多字号与强调格式 | 成功 | 通过 | 2085.47 | 2155.19 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/28.stdout.json) |
| 29 | 同一 run 的中西文字体 | 成功 | 通过 | 2079.09 | 2149.41 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/29.stdout.json) |
| 30 | 同段多组中西文字体 | 成功 | 通过 | 2161.45 | 2231.61 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/30.stdout.json) |
| 31 | 连续段落对齐与段间距 | 成功 | 通过 | 2267.37 | 2338.26 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/31.stdout.json) |
| 32 | 六种行距模式连续使用 | 成功 | 通过 | 2233.38 | 2302.95 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/32.stdout.json) |
| 33 | 大小字号与行距交叉 | 成功 | 通过 | 2228.25 | 2298.56 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/33.stdout.json) |
| 34 | 首行、悬挂与左右缩进 | 成功 | 通过 | 2275.18 | 2345.79 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/34.stdout.json) |
| 35 | 多 Action 追加与格式恢复 | 成功 | 通过 | 2574.21 | 2645.06 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/35.stdout.json) |
| 36 | 替换完整格式段落并保留邻段 | 成功 | 通过 | 2617.74 | 2689.22 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/36.stdout.json) |
| 37 | 跨 run 查找与多处格式化替换 | 成功 | 通过 | 2799.61 | 2869.53 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/37.stdout.json) |
| 38 | 综合文字报告与定点改写 | 成功 | 通过 | 2795.17 | 2869.76 | [JSON](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/38.stdout.json) |

## 数据与人工审核

- [Task统计CSV](../../../build/evidence/word-three-stage-20260917-234351/tasks.csv)、[逐Action统计CSV](../../../build/evidence/word-three-stage-20260917-234351/actions.csv)、[逐桥接统计CSV](../../../build/evidence/word-three-stage-20260917-234351/bridges.csv)、[统计汇总JSON](../../../build/evidence/word-three-stage-20260917-234351/summary.json)
- [原始报告](../../../build/evidence/word-three-stage-20260917-234351/windows/report.json)、[完整trace](../../../build/evidence/word-three-stage-20260917-234351/windows/logs)、[请求](../../../build/evidence/word-three-stage-20260917-234351/windows/original-requests)、[产物](../../../build/evidence/word-three-stage-20260917-234351/windows/outputs)
- [第三阶段清单与每步耗时](WPS_WORD_TEXT_FORMAT_TEST_CHECKLIST.md)、[第一阶段清单](WPS_WORD_FUNCTIONAL_TEST_CHECKLIST.md)、[第二阶段清单](WPS_WORD_ADVANCED_TEST_CHECKLIST.md)

Windows目录：`C:/Users/yim/wps-three-stage-20260917-234351`。输出名为 `234351` 轮次的 `20260917-234351-output-编号.docx/pdf`；14保存到本例fixtures输入原路径。临时计划任务已移除，测试文档保留供人工粗审。

本轮开始前WPS已经运行，测试期间按序逐步增加打开文档，未重置环境。每例只测一次，没有冷/热启动对照或重复性能样本，以上是本轮观测，不是稳定SLA、P95或吞吐率。原生回读判断效果，未额外解析最终DOCX/PDF或用截图判定；人工视觉审核仍待进行。
