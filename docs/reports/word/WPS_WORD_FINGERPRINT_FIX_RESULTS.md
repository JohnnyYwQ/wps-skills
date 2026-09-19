# Word 指纹读取替换与实机验证

日期：2026-09-17。已完成替换，并在 Windows WPS 12.1.0.28505、交互桌面 Session 1 验证通过。

## 实现与 contract

页眉页脚的观察读取由 `HeaderFooter.Range.Text` 改为解析内存中的 `Document.WordOpenXML`，按节、primary/firstPage/evenPages 变体和继承关系取得文字。读取磁盘文件不能看到未保存编辑，因此本实现只使用实时文档 XML。实际编辑页眉页脚仍使用 Range。

新增 `src/main/resources/wps_skills/word/windows/word_story_xml.ps1`，供快照、inspectDocument 和指纹共用；既不隐式保存，也不回写 `Document.Saved`。XML 无效或节映射不一致时明确失败。

14 个 Action 的参数及结果 schema 保持不变；更新 openDocument、inspectDocument 的验证约束和示例约束，并重新生成 actions.json。独立包包含新资源，56 个 manifest 文件校验通过。包位置：[wps-word](../../../build/archive/runs/word-fingerprint-fix-20260917-223921/wps-word)。[打包检查](../../../build/evidence/word-fingerprint-fix-20260917-223921/package-verification.json)。

## 验证结果

| 层级 | 结果 | 覆盖 |
| --- | --- | --- |
| 本地 unittest | 307 项，305 通过、2 项因需要 Windows PowerShell 跳过 | 合同、Task、资源打包等回归 |
| Windows XML 解析断言 | 16 项通过 | 空故事、变体、继承、显式空覆盖、隐藏样式、非法 XML/引用等 |
| Windows 原生指纹 | 4 例通过 | 普通文档、三节六变体、富文本页眉、已有页眉的未保存修改 |
| 正式 Task Client | 5 例通过 | 三种素材打开并重复检查；编辑已有文档并另存；新建写入、修改页眉、保存和导出 PDF |

原生四例打开时和读取后均 `Saved=true`，重复指纹一致。已有页眉的文档在主动编辑后，指纹由 `cef6f59bae88777f676cf336b0985fd29b00988f93af5548950e0f15576691d5` 变为 `e6dfe98d3a62b45f34b4b4d8151fa2f1e40ee1c713b5346df992f95419e08f8d`；再次读取稳定，且保留 `Saved=false`。这证明指纹能反映未保存的内存修改，观察本身不将这些样例标脏。

富文本页眉包含制表符、换行、表格、域结果、隐藏文字和修订；本例 XML 读数与原生 Range.Text 规范化结果相同。三节案例验证六种页眉页脚及第三节继承第二节。

Task 测试逐项检查 document、steps、completion 中的 Action Response 和总 Task Response；额外断言 saved 状态、revision 稳定/变化、替换数量及页眉文本。不存在 Agent 调用或 Task 重试。

## 人工粗审预期

| Task 案例 | 自然语言效果 |
| --- | --- |
| plain | 打开普通文档，两次检查不改变正文或页眉页脚。 |
| multi | 打开三节文档，各节和不同页型的页眉页脚保持原样。 |
| rich | 打开富文本页眉文档，检查后内容及排版保持原样。 |
| edit-save | 正文“甲”替换为“甲修改”，另存新文件成功。 |
| new-header | 新文档正文为“指纹回归”，主类型页眉为“新页眉”，生成 DOCX 和 PDF。 |

人工视觉审核尚未进行。本次为指纹问题的定向回归，未重跑原清单全部 18 例，也未验证所有复杂 Word 内容和大型文档性能。

## 证据与可复用入口

- [本地测试日志](../../../build/evidence/word-fingerprint-fix-20260917-223921/unittest.log)
- [原生指纹报告](../../../build/evidence/word-fingerprint-fix-20260917-223921/windows/native-hashes-r2/report.json)
- [Task 报告](../../../build/evidence/word-fingerprint-fix-20260917-223921/windows/native-tasks/report.json)，同目录保存请求、逐 Action 响应、Task 响应和输出文件。
- [原生测试脚本](../../../src/test/resources/wps_skills/word/fingerprint_acceptance.ps1)、[Task 测试脚本](../../../src/test/python/tests/word/fingerprint_acceptance.py)、[解析测试](../../../src/test/resources/wps_skills/word/story_xml_verification.ps1)。素材归档于证据目录 windows/fixtures。

首次原生测试的脚本加载器误把类构造器当成顶层函数，尚未进入文档操作即停止；修正测试加载器后完整通过。原始失败记录保留在 windows/native-done.json，成功记录在 windows/native-done-r2.json。通过后的测试脚本增加每轮唯一文件名前缀，避免 WPS 对已打开同名文档的冲突；生产代码未再变动。本次临时计划任务已清理，测试文档保留供人工查看。


后续验收更新（2026-09-17 22:56）：原 18 项功能清单已全量重跑，18 个 Task、80 个 Action 响应全部满足预期；人工粗审待进行。见 [最新实机结果](WPS_WORD_FUNCTIONAL_TEST_RESULTS.md)。上文未重跑的描述为当时阶段状态。
