# WorkBuddy Q4 第二轮：补齐 Task 结构说明后复测

2026-09-18，WorkBuddy 5.5.6，界面显示快速（Deepseek-V4.1-Flash）。新会话、新工作目录、新文件路径、新随机复核码；沿用用户已审核的提示词语义，未提供正文或预制 JSON。Agent 工作目录为 `C:/Users/yim/WorkBuddy/2026-09-18-11-36-46`。

**本轮内容目标及 Task JSON 有效性通过：提交2次、受理2次、成功2次；实际执行6个Action全部成功。没有参数拒绝或重试，没有读取底层源码，没有改动Skill包。** 这支持“先补齐SKILL.md再测”的诊断方向；单次成功不能将上一轮所有行为都归为文档全责，也未触发失败分支，不能据此宣称不重试规则已验证通过。

## 本轮唯一包变更

修改 `docs/drafts/wps-word/SKILL.md`，明确：document和completion条目只有address、params，不填id；steps条目才含id；address只含app、action。自检清单补充该区别，新增“打开、追加正文、保存”完整示例。其余行为规则、actions.json、runtime均保持不变，manifest只有SKILL.md哈希变化。

两个完整示例均通过当前runtime结构校验和Action参数预检。通用quick_validate因本地缺少PyYAML未运行成功，已手动检查frontmatter、引用路径，并核对包内全部manifest哈希。Windows安装前保留pre-draft-backup；真实Skill工具返回中确认加载到新增说明。测试结束后包内所有manifest文件哈希一致。

## 请求与响应

| Task | Action顺序 | Task结果 | Action结果 |
| --- | --- | --- | --- |
| read-01 | openDocument → inspectDocument | completed / succeeded | 2/2 succeeded |
| edit-01 | openDocument → replaceContent → inspectDocument → save | completed / succeeded | 4/4 succeeded |

两个Task均cleanup succeeded，输入文件已durably_admitted并消费。第一份只读响应完整包含正文（truncated=false）及随机复核码，直接通过Bash stdout返回。Agent随后生成同一路径的修改请求，使用query重新定位文字，没有跨Task使用旧range/revision。第二份completion正确只填address、params。

自然语言效果：选择状态为“通过”且费用最低的银杉方案，将采购结论改为：

> 采购结论：选择银杉方案，由负责人林澈负责，费用7400元，交付周期9天。

保存回原文件，其他六段正文保持原样。测试侧比较下载结果的DOCX正文，通过段落数、其余段落原文、结论四项事实和复核码检查。Agent最终复述新复核码 `Q4-99585feef8`，证明本次确实取得了新正文。未进行全版式人工验收。

## 耗时与额外行为

点击发送至最终答复 **79.363秒**；修改Task完整响应返回于发送后 **57.556秒**；之后收尾 **21.807秒**。已受理Task内部request.total合计 **10.687秒**。

| Task | 内部request.total | Action execute | 宿主命令往返 |
| --- | --- | --- | --- |
| 只读 | 4.486s | openDocument 2.197s；inspectDocument 0.871s | 5.075s |
| 修改保存 | 6.201s | openDocument 2.071s；replaceContent 0.793s；额外inspectDocument 0.814s；save 0.906s | 6.780s |

18次桥接往返合计8.038秒，匹配54条原生阶段记录，原生解析/执行/输出阶段合计6.148秒。差额1.890秒包含进程启动、调度、传输与日志等，不能当成纯通信。匹配采用bridgeRequestId+operation，嵌套耗时不可相加。

本轮仍有以下可观察偏差，未在中途追加提示或改规则：

- Agent额外安排read_back，只为比较改后正文，未用于后续编辑定位。草案已有“成功不额外回读验证”的规则，此处仍未遵守，增加约0.814秒执行成本。
- 定义查询不够直接：先Read整个actions.json，输出过大只返回预览；后又打印带缩进的大schema，再尝试提取字段遇到oneOf结构，最后导出所需定义读取。前置探查还受到宿主ls/dirname缺失及PowerShell无stdout影响。上述均非Task执行失败。
- 保存完成后调用Bash写工作日志（被宿主安全检查拒绝，返回理由为“Invoking PowerShell from Bash bypasses PowerShell security checks”；该命令实际是Python -c，是否误判未进一步诊断），再Glob、Write写入当前工作目录的记忆文件、present_files展示文档、生成最终答复。这些是约22秒收尾的可观察活动，没有新增Word Task。
- 最终答复有一处事实表述错误：称“未单独指定中文字体”，但实际修改JSON明确填写eastAsiaFontFamily=宋体。文档操作成功，答复仍需要精简并与请求/响应一致。

因此本轮可以确认“保存项JSON说明缺口已修正且本例不再报错”，尚不能把整个Agent行为都判为符合草案。首次失败后的停止行为本轮没有覆盖。模型随机性、宿主全局记忆等未受控，不能据两轮就证明唯一因果。

## 证据与状态

- [修正版草案](../../drafts/wps-word/SKILL.md)
- [独立请求和完整响应](../../../build/evidence/workbuddy-q4-r2-20260918-113531/submissions)
- [公开工具调用](../../../build/evidence/workbuddy-q4-r2-20260918-113531/actual/events.json)
- [最终答复](../../../build/evidence/workbuddy-q4-r2-20260918-113531/actual/agent-final.md)
- [最终界面](../../../build/evidence/workbuddy-q4-r2-20260918-113531/actual/observed-window-0.png)
- [结果DOCX](../../../build/evidence/workbuddy-q4-r2-20260918-113531/actual/q4-r2-20260918-113531.docx)
- [耗时数据](../../../build/evidence/workbuddy-q4-r2-20260918-113531/metrics.json)
- [测试后包校验](../../../build/evidence/workbuddy-q4-r2-20260918-113531/actual/package-after.json)

Windows当前保留修正版试用包，本地正式生产Skill未替换。临时计划任务已清理；文档、日志、备份保留。前一轮结果见 [第一轮报告](WPS_WORD_Q4_WORKBUDDY_RESULTS.md)。
