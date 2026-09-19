# openDocument 保存状态问题复现

日期：2026-09-17。已复现，定位到指纹读取阶段；未实施产品修复。

## 最小复现

Windows 交互桌面 Session 1；使用本轮原始 03.docx 的字节一致副本，每次使用全新文件名，避免复用已打开对象。文件 SHA-256：`f05d5ced0ddea16ca52180ad043af7609010241fe3d7286ad69f819967809e63`。

请求只包含 openDocument，`steps: []`、`completion: []`。不编辑、不保存，不设置 includeExistingChanges。

```json
{"app":"word","document":{"address":{"app":"word","action":"openDocument"},"params":{"path":"C:/Users/yim/word-open-state-20260917-220859/fresh-open-220859-a.docx"}},"steps":[],"completion":[]}
```

正式包运行两份独立副本：Task 与 openDocument 均 succeeded，但返回 `documentState.persistenceState == "modified"`，预期 saved 的断言两次失败。

运行入口为 [run.py](../../../build/evidence/word-open-state-20260917-220859/run.py)，由普通权限 Interactive 计划任务在桌面执行。已执行命令：

```text
ssh win powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File C:\Users\yim\word-open-state-20260917-220859\launch.ps1
```

已有路径和请求不可重放；再次复现需新建独立目录及文件名。[正式包复现结果](../../../build/evidence/word-open-state-20260917-220859/windows/report.json)。

## 隔离调试包的阶段观测

仅在 `build/evidence/word-open-state-20260917-220859/debug-word` 副本加入五个只读状态记录点，标记为 DEBUG-open-state-220859。正式源码与原测试包未改变。该 debug 副本用于诊断，不用于发布或功能通过结论；其原 manifest 不代表插桩后的哈希。

同一份全新文件的结果：

| 观察位置 | Document.Saved |
| --- | --- |
| Documents.Open 刚返回 | true |
| Show-BoundDocument 之前 | true |
| Show-BoundDocument 之后 | true |
| Get-DocumentFingerprint 之前 | true |
| Get-DocumentFingerprint 之后 | false |

变化发生于 2026-09-17 14:11:08 UTC 附近，指纹调用前后约 181 ms。该 probe 最终 openDocument 返回 modified，与未插桩正式包的两次复现一致。

[阶段原始记录](../../../build/evidence/word-open-state-20260917-220859/windows/probe-stages.jsonl) · [调试包响应结果](../../../build/evidence/word-open-state-20260917-220859/windows/probe-report.json)

## 结论与限制

本次文件刚打开时确实是 Saved=true。此前“文件加载后初始状态不符合”“可能要换成原生保存素材”的归因不成立，不能通过放宽 saved 预期或 includeExistingChanges 绕过。

状态变化已定位到 openDocument 内部 `Get-DocumentFingerprint` 执行期间；尚未进一步定位该函数内哪个 COM 属性读取导致变化，也未修复。下一步应细分该函数及其调用的节、页眉页脚读取，找出有副作用的访问；用户本轮要求先复现，因此停在上述证据结论。

两项临时计划任务已删除。未保存、关闭本轮文档，未结束 WPS。调试代码仅保留在标明用途的证据目录。


## 2026-09-17 22:23：已定位具体 COM getter

调用链：`openDocument → Get-DocumentFingerprint → Get-SectionSnapshots → Get-StorySnapshot → $story.Range`。

具体触发代码为 [word_bridge.ps1:584](../../../src/main/resources/wps_skills/word/windows/word_bridge.ps1#L584)。在第一节 primary header 上，取得 Item、读取 Exists、计算 Link 均保持 Saved=true；执行 `$range = $story.Range` 后立即变为 false，尚未读取 Range.Text，也尚未释放该 Range。

### 脱离 Skill 的最小对照

同一原始素材的四份独立副本，各使用新文件名，在交互桌面直接调用 WPS COM；不调用指纹函数，不执行内容写入或保存。

| 对照 | 取 Item、读 Exists 后 | 取 Range 后 / 对照等待后 |
| --- | --- | --- |
| 不取 Range，等待 250ms | true | true |
| primary header Range | true | false |
| primary header Range，独立副本复核 | true | false |
| primary footer Range | true | false |

所有用例刚 Open 后均为 Saved=true，Exists 均返回 true。因此不能靠当前的 `if ($exists)` 避免该副作用。

最小复现的关键操作：

```powershell
$doc = $documents.Open($path)
# $doc.Saved == true
$story = $doc.Sections.Item(1).Headers.Item(1)
$exists = [bool]$story.Exists
# $exists == true，$doc.Saved 仍为 true
$range = $story.Range
# $doc.Saved == false；尚未读取 $range.Text
```

[完整最小脚本](../../../build/evidence/word-fingerprint-20260917-222131/minimal.ps1) · [最小对照原始结果](../../../build/evidence/word-fingerprint-20260917-222131/windows/minimal-report.json) · [逐行插桩记录](../../../build/evidence/word-fingerprint-20260917-222131/windows/probe-stages.jsonl)

已执行入口：`ssh win powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File C:\Users\yim\word-fingerprint-20260917-222131\launch-minimal.ps1`。

### 当前结论

在本机 WPS 12.1.0.28505 与当前素材上，读取页眉或页脚的 Range getter 就能改变 Document.Saved。不是我们显式设置 Saved=false，也不是正文编辑或保存步骤造成的。我们的指纹代码把该 COM 访问当成纯读取，导致获取文档改变其保存标记，再触发 Task 原有修改保护。

这里只证明保存标记发生变化，未证明正文或页眉内容发生实际改写；也尚未证明所有文档、所有 WPS 版本都有此行为。WPS 是否在内部创建/初始化 story，仍只是可能解释，不作为已验证根因。

本轮目标为诊断，尚未修复正式代码。正式 bridge 与测试前独立包哈希一致。两项临时计划任务已移除；调试插桩只保留在证据目录 debug-word 中，不属于可发布包。


## 替代方式研究

后续已验证 Document.WordOpenXML 在本机可读取页眉页脚且保持 Saved，并可读取未保存内存修改；参见 [候选实测报告](WPS_WORD_HEADER_READ_ALTERNATIVES.md)。尚未接入正式代码。


## 2026-09-17 后续：正式替换与定向验收完成

已将观察时的页眉页脚 Range 读取替换为 live Document.WordOpenXML，并同步 contract 约束、生成的 actions.json、资源打包和测试。Windows 原生指纹 4 例、正式 Task 5 例通过；读取保持保存状态，未保存页眉编辑能改变指纹。上文“尚未替换”等描述保留为此前阶段记录；当前结果、包和证据见 [指纹修复验证](WPS_WORD_FINGERPRINT_FIX_RESULTS.md)。此次不代表原 18 项清单或迁移计划全部验收完成。
