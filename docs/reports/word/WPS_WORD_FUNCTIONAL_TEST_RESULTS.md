# Word Skill 实机功能测试结果

日期：2026-09-17 22:56。指纹修复后重跑 **18 项，18 项全部通过**；共核对 **80 个 Action Response** 和 **18 个 Task Response**。人工粗审尚未进行。

使用 `build/archive/runs/word-fingerprint-fix-20260917-223921/wps-word`，Windows 交互桌面 Session 1，56 个包清单文件校验通过。按原清单手工 JSON 经正式 Task Client 执行，每项只提交一次。业务 JSON 和断言不变，仅替换本轮路径及唯一文件名；核对脚本同步清单已明确的非末段换行预期。未执行 Agent 或 Task 重试测试。

已有文件用例 03, 04, 05, 06, 07, 14, 15, 16, 17 打开时全部报告 saved，之前由指纹观察造成的 modified 及后续保护拦截未再出现。所有 Action outcome/state 成功且无错误，Task completed/succeeded、stop=null；各例额外业务断言通过。详细动作、预期和完整响应见 [测试清单](WPS_WORD_FUNCTIONAL_TEST_CHECKLIST.md)。

## 逐项结果与人工粗审

| 编号 | 用例 | 响应结果 | 自然语言效果 | 人工粗审 |
| --- | --- | --- | --- | --- |
| 01 | 新建文字，不保存 | 通过 | WPS 中出现一份未保存的新文档：标题“项目周报”，下方依次是中文正文和英文正文。 | 未审阅 |
| 02 | 文字与段落格式 | 通过 | 一份格式样张：蓝色、18 磅、加粗、居中的一级标题；正文中文宋体、西文 Times New Roman，12 磅，首行缩进 24 磅，1.5 倍行距，段前后各 6 磅。 | 未审阅 |
| 03 | 在指定段落前插入 | 通过 | 原来的“甲、乙、丙”三段变成“甲、新增段、乙、丙”，原段落保持原样。 | 未审阅 |
| 04 | 读取及查找 | 通过 | 文档保持原样。返回的正文是“Alpha alpha Alphabet”和“项目：待确认”；区分大小写的完整词 Alpha 有 1 处，不区分大小写有 2 处，“不存在”有 0 处。 | 未审阅 |
| 05 | 批量替换文字 | 通过 | “设计：待确认”“开发：待确认”“测试：待确认”三段中的“待确认”全部变成“已确认”。 | 未审阅 |
| 06 | 替换指定完整段落 | 通过 | 三段“开头、旧说明、结尾”变成“开头、新说明、结尾”；中间段落加粗、居中。 | 未审阅 |
| 07 | 删除指定完整段落 | 通过 | 三段“保留甲、删除我、保留乙”删除中间段后，仅保留甲、乙两段。 | 未审阅 |
| 08 | 插入表格 | 通过 | “进度表”下面是一张 4 行 3 列的表格，首行为表头；表格下方有独立正文“以上为本周进度。”。 | 未审阅 |
| 09 | 插入图片 | 通过 | 文字“图片示例”之后嵌入蓝底白字 TEST 图片，宽 2 英寸、高 1 英寸，替代文本为“蓝底白字 TEST 图片”。 | 未审阅 |
| 10 | 页面方向与页边距 | 通过 | 页面横向，四边页边距均为 1 英寸；正文为“横向页面样张”。 | 未审阅 |
| 11 | 页眉页脚 | 通过 | 两页文档，第一页和第二页各有相应正文；每页页眉“项目报告”、页脚“内部资料”。 | 未审阅 |
| 12 | 插入分页符 | 通过 | “第一部分”在第一页，“第二部分”从下一页开始。 | 未审阅 |
| 13 | 分节与第二节横向 | 通过 | 第一节纵向；第二节另起一页且横向，四边页边距为 1 英寸。 | 未审阅 |
| 14 | 编辑已有文件并保存 | 通过 | 原文件“版本：旧版本”变成“版本：新版本”，后面增加“新增说明”，保存到原路径。 | 未审阅 |
| 15 | 编辑已有文件后另存 | 通过 | 生成新的 15.docx，正文为“版本：新版本”；原输入文件仍是“版本：旧版本”。 | 未审阅 |
| 16 | 编辑已有文件，不保存 | 通过 | WPS 打开的文档显示“版本：新版本”，处于有未保存修改的状态；磁盘原文件仍是旧版本。 | 未审阅 |
| 17 | 修改后仅导出 PDF | 通过 | 17.pdf 内容为“版本：新版本”；WPS 文档保留未保存修改，磁盘 DOCX 仍是旧版本。 | 未审阅 |
| 18 | 完整报告：DOCX 与 PDF | 通过 | 一份两页项目周报：第一页标题、正文和进度表；第二页“附图说明”和 TEST 图片；两页都有页眉页脚。同时交付 DOCX 和 PDF。 | 未审阅 |

## 文件与原始证据

Windows 本轮目录：`C:/Users/yim/wps-functional-rerun-20260917-225604`。素材名为 `fixtures/r20260917225604-fixtures-编号.docx`；输出名为 `outputs/r20260917225604-outputs-编号.docx` 或 `.pdf`。14 在素材原路径保存；01 未保存的新文档、16/17 的未保存修改保留在 WPS 中供人工查看。

- [汇总报告](../../../build/evidence/word-functional-rerun-20260917-225604/windows/report.json)
- [原始响应目录](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses)，包括每个 Task 中完整 Action 错误/数据字段。
- [原请求](../../../build/evidence/word-functional-rerun-20260917-225604/windows/original-requests)、[输出文件](../../../build/evidence/word-functional-rerun-20260917-225604/windows/outputs)、[素材元数据](../../../build/evidence/word-functional-rerun-20260917-225604/windows/fixture-metadata.json)
- [本轮执行脚本](../../../build/evidence/word-functional-rerun-20260917-225604/windows/harness/run.py)、[响应核对脚本](../../../build/evidence/word-functional-rerun-20260917-225604/windows/harness/verify.py)
- [首轮结果存档](../../../build/evidence/word-functional-rerun-20260917-225604/results-before-rerun.md)、[首轮清单存档](../../../build/evidence/word-functional-rerun-20260917-225604/checklist-before-rerun.md)

启动准备时压缩包遗漏空 responses/outputs 目录，脚本在提交任何 Task 前停止；补齐目录后开始正式执行，未重放 Action 或 Task。准备错误原样保留在 [bootstrap-report.json](../../../build/evidence/word-functional-rerun-20260917-225604/windows/bootstrap-report.json)。

本轮按约定信任 Action 内部回读，未额外解析最终 DOCX/PDF 或做截图验收；人工审核结果独立记录。


最新统一复测：三阶段38例已使用正式trace包重新执行，结果和Task/Action/通信耗时见 [三阶段统一报告](WPS_WORD_THREE_STAGE_RESULTS.md)。本文件上文保留为历史轮次记录。
