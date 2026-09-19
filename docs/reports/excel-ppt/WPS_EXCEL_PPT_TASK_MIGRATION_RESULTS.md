# Excel / PPT 共用 Word Task 机制：实现与验证结果

2026-09-18。Excel 与 PPT 已接入 Word 的当前 Task 工作流，并分别生成可独立安装的 `wps-excel`、`wps-ppt`。本轮完成了源码、Skill、契约查询、执行入口、原生操作接入、测试和打包；未发布新的 GitHub Release。已发布的 Word 2026.09.18 包及分支未改动。

## 复用与应用差异

| 共用一份实现 | 各应用独立维护 |
| --- | --- |
| 版本无关的 document / steps / completion 请求与 Task Response | create/open Action、文档模型和应用注册 |
| 参数校验、前序结果引用、保存与 PDF 的执行顺序 | Action 名称、参数、结果契约和语义验证 |
| 输入路径提交身份、持久回执、重复提交只读查询 | Word 范围；Excel A1 区域和工作表；PPT 幻灯片及形状 ID |
| 失败/不确定停止、禁止重放、已有修改保存保护 | 各应用观察范围和 token 生成/核对 |
| Task/Action 执行器、Windows 资源装配、桥接、租约及清理 | 原生读写、原生回读验证、保存和导出的具体调用 |
| 请求、Action、桥接和原生阶段的计时 | 应用能力边界 |
| schema 导出与 `$ref` 公共定义去重、按名称查询脚本 | 各 Skill 的用途表、定位规则及参数说明 |

三个 Skill 的使用顺序一致：先按用途编排 → 查询已选 Action 定义 → 填写 JSON → 提交 → 读 Task Response → 自然语言交付。没有恢复旧 Session 接口，没有让 Agent 编写控制脚本，也没有增加 Agent 自动重试。

Excel 34 个 Action 覆盖新建/打开、工作簿/工作表读取、区域读写、公式与计算、格式、工作表增删复制移动、行列增删、合并、排序筛选、查找替换、复制及 XLSX/PDF。

PPT 37 个 Action 覆盖新建/打开、幻灯片增删复制排序、文本与中西文字体、段落、形状布局与外观、图片、表格、备注、页面设置及 PPTX/PDF/PNG。`exportSlideImage` 位于 `steps`，可引用本 Task 新建幻灯片的结果；保存和整份 PDF 仍在 `completion`。

## 最终实机结果

环境：Windows 已登录桌面、真实 WPS、Python 3.13.12。使用隔离安装包和预编排 JSON，经包内正式 CLI 提交；未使用 Agent、旧 Session 或生产目录中的已安装 Skill。测试根目录：

`C:/Users/yim/wps-task-migration-20260918-final`

| 应用 | 正向 Task | 正向 Action | 异常保护用例 | 成功执行覆盖的不同 Action |
| --- | --- | --- | --- | --- |
| Excel | 3/3 成功 | 64/64 成功 | 6/6 符合预期 | 34/34 |
| PPT | 3/3 成功 | 61/61 成功 | 6/6 符合预期 | 37/37 |
| Word 共享机制回归 | 1/1 成功 | 4/4 成功 | 本轮未新增 | 4 个 |
| 合计 | 7/7 成功 | 129/129 成功 | 12/12 符合预期 | 不跨应用合并 Action 名称 |

**19/19 是用例符合预期率；不是 19/19 Task 返回 succeeded。** 12 个异常用例应返回 failed，其中每个应用 5 个是明确的 Action 失败、1 个是打开文档成功后命中已有未保存修改保护。全批次共 145 个 Action 成功、10 个 Action 按预期失败、8 个未执行，无 unknown。

全部 19 个 Task 的原输入路径回执查询与提交结果一致，输入文件均在持久接收后被消费。所有 request / Task / Action 日志的应用标识与 Task Response 相符。最终包与 Windows 实测包逐文件 SHA-256 一致。

### 每个应用的用例

| 用例 | 成功条件 |
| --- | --- |
| 新建、完整内容操作链、保存、PDF | Excel 59 个 / PPT 56 个 Action 连续成功；PPT 同时覆盖 PNG；产物为有效文件，保存包可读取。 |
| 打开既有文件并保存 | 精确打开本轮产物，读取和 save 成功。 |
| 使用过期 token 修改 | Excel STALE_RANGE / PPT STALE_CONTENT；后续内容步骤 not_executed。 |
| 未授权保存既有未保存修改 | TASK_EXISTING_CHANGES_CONFIRMATION_REQUIRED；内容与保存均未执行。 |
| 授权包含既有修改后保存 | includeExistingChanges=true 后独立请求成功。 |
| 目标文件正由另一文档打开 | OUTPUT_IN_USE，原文件哈希不变。 |
| 目标文件已存在但未打开 | OUTPUT_ALREADY_EXISTS，原文件哈希不变。 |
| 输出父目录不存在 | OUTPUT_PARENT_NOT_FOUND，不创建缺失目录。 |
| 输入文件不存在 | DOCUMENT_NOT_FOUND，不转为新建。 |

## 执行时间

以下是最终轮次的完整操作链，不含 Agent 编排时间，不代表日常短任务平均值：

| 请求 | Action 数 | CLI 进程墙钟 | JSON 输入至响应写出 | Action 执行累计 | 桥接往返累计 |
| --- | --- | --- | --- | --- | --- |
| Excel 完整操作链 | 59 | 50.38 秒 | 50.30 秒 | 49.74 秒 | 49.67 秒 |
| PPT 完整操作链 | 56 | 5.25 秒 | 5.16 秒 | 4.80 秒 | 4.74 秒 |
| Word 简单内容 + DOCX/PDF | 4 | 1.84 秒 | 1.76 秒 | 1.55 秒 | 1.56 秒 |

桥接往返包括原生 WPS 工作及等待，不能称为纯通信时间；它还包括绑定探测与清理，累计值不要求小于内容 Action 累计时间。Excel 的逐单元格读取、写入、复制和工作表观察明显更贵，迁移成功并不表示三个应用具有相同执行速度。

分别记录 `input.read/decode`、Task admission/preflight/execution/cleanup、每个 `action.execute`、bridge encode/write/wait/decode、native decode/operation/response-write。完整逐 Action 耗时见 [action-timing.csv](../../../build/archive/runs/excel-ppt-migration/verified-evidence/action-timing.csv)，分阶段结果见 [summary.json](../../../build/archive/runs/excel-ppt-migration/verified-evidence/summary.json)。

## 本地验证与安装包

197 项当前 unittest：194 通过、3 项 Windows 专项在 macOS 跳过。新增检查覆盖两应用请求预编排、非法生命周期与前向引用拒绝、失败停止、无重放回执、既有修改保护、三应用资源工厂、日志应用归属、schema 注释不放松校验、全部参数示例、handler 覆盖和安装包迁移后使用。

两应用全部查询 schema 展开 `$ref` 后，与运行时的 parameters/result/examples 完全相等。每份包可在没有其他应用 Python 包的环境下独立导入与查询；拒绝以另一个 app 调用。两份 SKILL.md 均通过 Skill 校验脚本。

- [Excel 安装目录](../../../build/archive/runs/excel-ppt-migration/final/skills/wps-excel)：90 文件，336,284 字节。
- [Excel ZIP](../../../build/archive/runs/excel-ppt-migration/final/wps-excel.zip)
- [PPT 安装目录](../../../build/archive/runs/excel-ppt-migration/final/skills/wps-ppt)：103 文件，353,466 字节。
- [PPT ZIP](../../../build/archive/runs/excel-ppt-migration/final/wps-ppt.zip)
- [完整 Task 回执与预期](../../../build/archive/runs/excel-ppt-migration/verified-evidence/report.json)
- [本地测试日志](../../../build/archive/loose-files/excel-ppt-migration-tests.log)
- [Windows 证据归档](../../../build/archive/runs/excel-ppt-migration/verified-evidence.zip)

安装时复制完整 `wps-excel` / `wps-ppt` 目录。构建入口分别为 `scripts/build/excel.py`、`scripts/build/ppt.py`，原生测试入口为 `scripts/validate/applications.py`。

## 迁移中发现并处理的问题

1. Task Client、请求编译、回执初始化、CLI 和构建原本只允许 Word。抽出共享实现，通过应用自有契约与工厂接入两应用；未复制三套生命周期。
2. token 的来源不是统一的“上一条返回”：PPT 列表、幻灯片、样式、表格、备注、设置分别有其观察 token。Skill 与参数 schema 明确来源，并以真实结果引用验证连续编排。
3. 首轮重名测试把“已存在”与“正打开”混为一例。拆成两例分别核对准确错误码；随后重跑因上一轮打开文件同名而被保护拦截，测试文件改为每轮唯一名称。未关闭用户文档或放宽保护。
4. 原 Word 计时日志有硬编码的应用名。修正为 Task 所属应用，补充测试并再次完成原生全量回归。

## 当前能力边界

- 本轮是直接 JSON 的原生执行验证，**不是 Excel/PPT 的 Agent 端到端测评**；尚未测其自然语言任务完成率，尚无人工作品粗审结论。
- Excel 每次区域操作最多 1000 格；保存/导出验证最多 20 张工作表，每张已用区域最多 1000 格；公式采用当前允许的英文函数及同工作表引用。
- PPT 最多 200 页、每页 100 个顶层形状；表格最多 20 行、10 列、100 格。此次不是对任意复杂模板或 Office 全部功能的兼容性证明。
- Excel/PPT 当前保存与导出使用 failIfExists，未把 Word 的 renameIfExists / PDF replaceExisting 自动推广为两应用能力。
- Word 的原有广泛测试证据继续归属 Word；本轮仅额外验证共享机制变更后的一个原生 Word Task。
