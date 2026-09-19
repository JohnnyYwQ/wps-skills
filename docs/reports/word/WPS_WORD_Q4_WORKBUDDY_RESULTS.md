# WorkBuddy Q4 草案实测结果

2026-09-18，Windows WorkBuddy 5.5.6，界面显示“快速（Deepseek-V4.1-Flash）”。真实 Agent 自行读取 Skill、查 Action 定义、生成 JSON 并调用 Task Client；未提供正文、答案或预制请求。提示词经用户审核后发送，仅运行这一轮。

**结论：本例跨 Task 使用正文成立，最终文档正确；草案整体验收未通过。** 首次修改请求被拒绝，Agent 自行重试，还读取底层代码并修改了安装的 SKILL.md。最终成功不能抵消这些偏差。

## 实际请求与响应

| 提交 | 编排 | Task Response | Action Response |
| --- | --- | --- | --- |
| 1：read-1 | openDocument → inspectDocument；completion=[] | completed / succeeded；cleanup succeeded | 2 个 Action 均 succeeded；正文完整、truncated=false |
| 2：edit-2 | openDocument → replaceContent → save | rejected / failed；INVALID_TASK_REQUEST；validation | 尚未执行 Action；replace_conclusion 为 not_executed、response=null |
| 3：edit-3（Agent 自行重试） | openDocument → replaceContent → save | completed / succeeded；cleanup succeeded | 3 个 Action 均 succeeded |

第二次提交给 completion[0] 多填了 `"id":"save"`。执行入口要求该项只有 address、params，返回 `Expected object fields: address, params`。错误发生在 Task 结构校验，taskId=null、recordPath=null、输入保留且 not_admitted，因此没有已执行但失败的 Action。拒绝响应从实际 Bash stdout 提取保存，不能仅统计落盘回执，否则会漏掉这次失败。

提交口径成功率 **2/3（66.7%）**；已受理 Task 为 **2/2**；实际执行 Action 为 **5/5**。首次修改提交未成功，严格“不重试”流程应在此报告停止。

## Q4 的证据

- Skill 工具返回的正文与本次安装草案一致，包含先读后改、参数拒绝也不重试的规则。
- 第一次 Task 的完整响应直接出现在 Bash stdout，正文含三项候选、负责人和随机复核码。Agent 没有依赖额外文件读取取得正文。
- 第一次 Task 已释放文档资源，桥接进程 35968 已退出；之后 Agent 才根据返回正文生成修改请求。
- 修改请求已正确选出银杉方案、负责人林澈、7400元、9天，说明首次修改提交在语义上已经衔接成功。它通过 query 定位目标文字，没有沿用前一 Task 的 range/revision。
- 真正受理的修改 Task 使用新的输入路径、Task ID 和桥接进程 29864，重新获取同一文件，最终保存成功。两次入口均完成返回。
- 最终回答正确复述复核码 Q4-9ba2c90c51。

本例证明当前宿主中，执行资源清理后 Agent 仍能使用已返回的短文档内容。没有人为销毁整个 WorkBuddy 对话或工作目录，也未测试长正文截断、上下文压缩、新对话、重启或未保存新文档，不能外推这些情况。

## 自然语言效果

采购结论变为：

> 采购结论：经比选，选用银杉方案，由负责人林澈牵头执行，费用7400元，交付周期9天。

其余六段文字和随机复核码保持原样，保存回原文件。测试侧对下载结果的 DOCX 正文作了逐段比较，以上内容条件通过；未据此声称所有版式属性逐一验证。WorkBuddy 最终界面展示了修改后的文档。

## 草案暴露的问题

1. **Task 外层结构说明不够完整。** Action schema 只定义 params/result，不能补全 document、steps、completion 三处容器结构。草案只有 completion=[] 的整例，没有显式说明 completion 项不得填 id。Agent 为此搜索并读取 runtime 源码，仍首次填错。
2. **不重试规则未被遵守。** 草案已经明确涵盖参数拒绝。Agent 自行以“没有副作用”为理由增加例外，改用新路径重提。不能说草案没有这条规则；只能说这次没有约束住实际行为。
3. **Agent 自行改了 Skill。** 保存完成后调用 Edit 修改安装目录 SKILL.md，补充 completion 说明，并写入本次工作目录的记忆文件。这不属于用户文档目标。仅从本次公开记录无法确认宿主其他指令对其行为的影响。
4. **宿主 Bash 有环境杂音。** ls/mkdir/dirname 不可用，产生额外探查和错误；Python 入口及两次已受理 Task 正常工作。此类错误不是 WPS Action 失败。

下一版应明确给出三处请求结构及包含 save 的完整例子，说明遇到拒绝立即交付失败结果（包括未受理、无副作用的情况），并限定文档处理期间不修改 Skill/contract/runtime。此次未替换本地草案或正式源码，也没有改版重测。

## 耗时

从点击发送到最后一条公开回答约 **105.65 秒**，包含 Agent 编排、代码查找、失败提交、重试和额外收尾。

| 已受理 Task | CLI 内部 request.total | 各 action.execute | 桥接往返合计（含远端执行） |
| --- | --- | --- | --- |
| 只读 | 8.222 s | openDocument 5.446 s；inspectDocument 1.000 s | 7 次，6.711 s |
| 修改并保存 | 5.466 s | openDocument 2.128 s；replaceContent 0.813 s；save 0.883 s | 9 次，4.044 s |

两个已受理请求内部合计13.688秒；被拒绝调用的宿主工具往返约0.881秒，该值与内部计时边界不同。桥接16次与原生48个阶段按 bridgeRequestId+operation 完整匹配。往返总10.754秒，原生解析、执行、写出阶段总8.696秒，差额2.058秒包含启动、调度、传输与日志等，**不是纯通信时间**。以上为单次观察，不作为性能基准。

## 证据与安装状态

- [审核过的提示词](../../../build/evidence/workbuddy-q4-20260918-110849/actual/prompt.txt)
- [三个请求及对应响应](../../../build/evidence/workbuddy-q4-20260918-110849/submissions)
- [公开调用记录](../../../build/evidence/workbuddy-q4-20260918-110849/actual/events.json)
- [Agent 最终答复](../../../build/evidence/workbuddy-q4-20260918-110849/actual/agent-final.md)
- [最终界面](../../../build/evidence/workbuddy-q4-20260918-110849/actual/final-window-0.png)
- [输出文档](../../../build/evidence/workbuddy-q4-20260918-110849/actual/q4-20260918-110849.docx)
- [耗时数据](../../../build/evidence/workbuddy-q4-20260918-110849/metrics.json)
- [Agent 改过的 Skill 快照](../../../build/evidence/workbuddy-q4-20260918-110849/SKILL.after-agent.md)

Windows 安装位置仍是试用包；已保留 Agent 的 Skill 改动证据，并恢复至用户审核的原草案。旧安装备份仍在测试目录 installed-backup / pre-draft-backup。正式源码未替换；临时发送和观察计划任务已清理。Agent 在本次 WorkBuddy 工作目录中写入的记忆文件保留，因此后续独立测试应使用新的工作目录，不能宣称该会话仍是首次接触。
