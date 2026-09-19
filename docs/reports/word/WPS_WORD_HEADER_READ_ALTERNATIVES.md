# 页眉页脚无副作用读取候选

日期：2026-09-17。结论：本机已找到可行候选 `Document.WordOpenXML`，尚未修改正式代码。Windows WPS 12.1.0.28505，交互桌面 Session 1。

## 当前建议

优先考虑从 **live Document.WordOpenXML** 获取页眉页脚内容，按节引用、变体和继承关系解析。它不是重新读取磁盘 DOCX，因此能包含未保存的内存修改。官方文档将该入口定义为整个文档的 Flat XML：[Word Document.WordOpenXML](https://learn.microsoft.com/en-us/office/vba/api/word.document.wordopenxml)、[WPS 对应 API](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wps/Document/member/WordOpenXML)。WPS 官方页面是 JS API 文档，本报告的 COM 可行性以本机实测为依据。

可行候选不等于完成替换：当前测试覆盖普通文本、三节六变体、链接继承、未保存页眉修改。复杂域、修订、文本框、表格等内容如何保持与现有 Range.Text 规范化一致，以及大型文档成本，还没有完成生产验收。

## 实机比较

| 路径 | 结果 | 结论 |
| --- | --- | --- |
| Document.WordOpenXML | 无页眉、已有页眉页脚、三节样例读取前后均 Saved=true；未保存修改后的新文字可读且 Saved=false 保持 | 当前最有希望的替代入口 |
| Document.StoryRanges + NextStoryRange | 读取已有文字不改变 Saved，能看到未保存修改；三节例仅返回两个独立故事节点，Information(2) 对节点均返回 1 | 有效读取候选，但缺少可靠节映射；不可直接用链序号替代节号 |
| HeaderFooter.IsEmpty 后决定是否取 Range | 空白 primary 页眉的 Exists=true、IsEmpty=false；取 Range 仍使 Saved=false | 无法作为保护条件 |
| Document.Content.WordOpenXML | 当前样例也返回了页眉页脚，Saved 保持 | 不优先：官方只承诺范围所需 XML，不保证完整文档所有内容 |
| 解压磁盘 DOCX | 无法反映未保存内存修改 | 不适合作为 live 指纹内容来源，未作为候选执行 |

StoryRanges 的官方说明：[StoryRanges](https://learn.microsoft.com/en-us/office/vba/api/word.storyranges)、[NextStoryRange](https://learn.microsoft.com/en-us/office/vba/api/word.range.nextstoryrange)。本机对不存在的类型返回 E_FAIL，不能把所有异常无条件当作空内容。

## XML 候选已验证的事实

1. 空白页眉页脚素材：读取前后 Saved=true；没有伪造页眉页脚文字。
2. 已有内容素材：准确读到“项目报告”和“内部资料”，Saved=true 保持。
3. 多节素材：第 1、2 节各有 primary/first/even 三种页眉与页脚，共 12 份独立文字；第 3 节链接第 2 节。按 document.xml 中 sectPr、r:id 与 relationships 解析，18 个 section/area/variant 对应的文字和继承断言全部通过。
4. 未保存内存修改：在专用测试副本将页眉改为 `LIVE_UNSAVED_HEADER_222854`，未保存；StoryRanges 和 Document.WordOpenXML 都读到新文字，读取前后 Saved=false。磁盘文件没有被用于得出新文字。

XML 映射原型仅针对上述纯文字样例验证关系与继承，不是完整生产文本解析器。不能直接拼接所有 w:t 或 hash 整篇 XML 就宣称等价于现有指纹。

## 对原问题的进一步限定

本轮在“已有实际页眉内容”的样例上读取 HeaderFooter.Range，没有改变 Saved；原先无相应内容的样例上才复现取 Range 变脏。因此目前应描述为“特定空页眉页脚状态下，该 getter 有副作用”，而非“所有 Range 读取都会变脏”。WPS 内部是否懒创建了 story，仍未被直接证实。

## 证据与测试边界

- [官方资料研究](../../../build/evidence/word-header-read-research-20260917.md)
- [候选读取原始报告](../../../build/evidence/word-header-read-20260917-222854/windows/report.json)
- [IsEmpty 独立复核](../../../build/evidence/word-header-read-20260917-222854/windows/isempty-report.json)
- [未保存内存读取报告](../../../build/evidence/word-header-read-20260917-222854/windows/live-report.json)
- [18 项节/变体映射断言](../../../build/evidence/word-header-read-20260917-222854/xml-mapping-verification.json)
- [映射验证原型](../../../build/evidence/word-header-read-20260917-222854/verify_xml_mapping.py)

首轮 IsEmpty 探针把 COM collection 放在 PowerShell if 表达式输出中，造成枚举后索引偏移；该组三条初始结果无效，不作为结论。随后改为分支内直接赋值，在全新文件上独立运行 isempty.ps1，结果以上述独立复核报告为准。其他候选路径不使用该赋值逻辑。

正式代码与正式包未改动；没有强制设置 Saved、保存或关闭文档。仅 live 探针在明确的独立测试副本中修改了一处页眉，以验证读取内存新值。三项临时计划任务已移除，脚本与文档留在独立证据目录。


## 2026-09-17 后续：正式替换与定向验收完成

已将观察时的页眉页脚 Range 读取替换为 live Document.WordOpenXML，并同步 contract 约束、生成的 actions.json、资源打包和测试。Windows 原生指纹 4 例、正式 Task 5 例通过；读取保持保存状态，未保存页眉编辑能改变指纹。上文“尚未替换”等描述保留为此前阶段记录；当前结果、包和证据见 [指纹修复验证](WPS_WORD_FINGERPRINT_FIX_RESULTS.md)。此次不代表原 18 项清单或迁移计划全部验收完成。
