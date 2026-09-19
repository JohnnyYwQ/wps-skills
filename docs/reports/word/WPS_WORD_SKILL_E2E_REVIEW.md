# 端到端测试前的 SKILL.md 检查

日期：2026-09-18。检查对象：src/main/resources/skills/wps-word/SKILL.md。当前仍为正式 14 Action 方案。

## 发现并修订

| 问题 | 修订 |
| --- | --- |
| 总述“继续、调整”和参数拒绝后“按错误修正”暗示自动重试 | 明确提交前自查一次；提交后无论拒绝、失败或 unknown 都报告并停止，用户要求继续才另行处理 |
| 查询要求 Bash，但代码块标为 powershell | 改为 text；查询和提交均明确 WorkBuddy 用 Bash 直接调用 Python |
| 查询输出截断时没有明确补齐要求 | 使用现有脚本分批查询补齐，不能凭缺失规则猜测 |
| 结果引用 data 前缀不够突出 | 明确 result 描述 Action Response 的 data，Task 结果引用必须以 data 开头 |
| remainingRange 描述可能被理解成可跨 Task 沿用 | 明确只在原 Task 有效，新 Task 重新定位；不拼接不同 revision 的快照 |
| 未保存新文档跨 Task 继续的边界不明确 | 新文档尽量一次编排，不为接续擅自保存 |
| 缺少提交前一次编排自查 | 检查目标、参数、引用、范围和交付，不另写校验脚本 |
| 有内部实现说明占用注意力 | 删除自动换名最大候选数、资源释放过程、清理结果字段讲解及内部错误码举例等 |
| 最终汇报可能扩张为重复讲解或额外展示脚本 | 明确简短自然语言，说明完成/未完成及实际路径，不复述 JSON 或另写展示脚本 |

## 保留的必要信息

- 14 个 Action 的用途和 document/steps/completion 放置规则。
- 输入 schema、结果 schema 和共享定义，以及两种 $ref 的区别。
- Task 范围/版本不能跨请求沿用、结果路径、步骤编号限制：这是正确编排所需的外部约束。
- 参数层级、保存意图、已有修改授权、UTF-8 请求文件、Bash/Python 调用方式：这是当前可用入口的必要使用说明。
- 查询回执的路径和 taskId、失败/unknown/not_executed 区别：这是如实报告结果所需信息。

没有向 Agent 引入 COM、handler、指纹算法、租约、进程或后端协议。

## 核对依据与验证

核对了 word/task/plan.py 的请求结构与生命周期顺序、client/task_client.py 的结果状态、cli/call.py 的命令参数和默认超时、ADR 0029 的 Task 边界。

- 8 项既有 schema 查询及 Word 包构建测试通过。
- SKILL.md 的全部独立 JSON 代码块经解析和当前 contract 校验；完整 Task 示例通过编排预检。
- 20 个端到端用例编号唯一且齐全。
- Frontmatter 与命令代码块做了本地检查。没有将缺少 PyYAML 时不可用的 Skill quick_validate 宣称为通过。

本轮修改仅为文档与测试方案；未修改 contract、Action、handler 或执行层。尚未构建/部署本轮 Windows 测试包，尚未运行 20 例。后续测试应使用本次修订后的正式 Skill 源构建，不使用旧草稿或拆分实验包。

## 第二次精简（根据用户复核）

此前修订仍混入历史故障防御性指令和重复说明，本次重新整理为五步：选择能力 → 查询 → 编排自查 → 提交读回执 → 交付。

- Bash/Python 调用方式只在开头说明一次；删除“能返回 stdout”、不使用 PowerShell、不直接 Read、不加 head/管道等重复禁止项。
- 删除自动追加编号的行为说明；只保留 Agent 必须选择的 overwritePolicy 语义及最终使用 artifact.path 的规则。
- 合并三部分 Action 表，删除重复的保存顺序、定义字段解释和提交说明。
- 保留 UTF-8 文件写入方式、data 前缀、范围有效期、参数规则、回执查询和不重试约定。
- 从 6,544 字符减少至 4,322 字符，约减少 34%。完整 Task 示例再次通过当前请求编排及参数预检；结果引用示例经过 JSON 解析和 data 前缀检查。

历史 PowerShell 问题的准确描述：WorkBuddy PowerShell 工具只返回退出码，没有回传 Python stdout；不是已经确认 Python 没有生成输出。该故障解释留在评审记录，不放进面向 Agent 的操作步骤。

## 第三次收敛（当前版本）

用户进一步明确：先按用途选择 Action/安排顺序，再查 schema 填参数；无需独立自查步骤，只消费 Task Response 判定结果。

已删除特定宿主名称、PATH/解释器回退说明、独立 Action 成功判定、自查步骤、taskId/recordPath 备用查询分支。保留一个按原输入路径取得 Task Response 的查询入口，沿用此前允许查询回执的约定。保存相关只留下 Agent 必须填写的意图/参数选择，不描述内部过程。JSON 示例再次通过当前 contract 预检。

以上取代本文前两轮中“提交前自查一次”等过时描述。正式源文档已更新，测试尚未启动。
