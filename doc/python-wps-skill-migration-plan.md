# WPS Skills Python 化迁移计划

## 1. 目标

将当前基于 Node.js、TypeScript 和 MCP 的 WPS 自动化仓库迁移为一个可由目标智能体加载的 `wps-office` Skill。

最终产物满足以下约束：

- 不依赖 Node.js、npm、MCP 协议或 MCP SDK；
- 目标智能体加载 `SKILL.md` 后调用技能目录内的 Python；
- Python 3.9 及以上，仅使用标准库；
- 保留现有 Excel、Word、PowerPoint、通用及跨应用能力；
- 不保留旧 MCP Tool 名称和 MCP 响应形式；
- Linux x86_64、Linux ARM64、Windows x86_64、Windows ARM64 使用同一套 Python 和 WPS JavaScript 加载项；
- 每次 Python 调用只执行一个 WPS Action，完成后退出；
- 首次使用时自动安装或更新用户级 WPS 加载项。

相关架构决定见 [`docs/adr/`](../docs/adr/)。

## 2. 最终边界

```text
用户请求
   ↓
目标智能体加载 SKILL.md
   ↓
读取相关 references
   ↓
执行 scripts/wps.py
   ↓ 127.0.0.1 临时轮询服务
WPS JavaScript Add-in
   ↓
WPS Office
```

迁移后不再存在：

- MCP Server；
- MCP Client 配置；
- TypeScript ToolRegistry；
- Node.js WpsClient；
- Node.js 轮询服务器；
- Windows PowerShell COM 执行路径；
- 分散的四个独立 WPS Skill。

保留并改造：

- `wps-claude-assistant/main.js` 及其 WPS 操作实现，作为统一加载项的基础；
- 已有 WPS Action 名称、参数和结果语义；
- 现有 Skill 中仍然有效的操作流程、安全规则和错误处理知识。

## 3. 最终技能目录

```text
wps-office/
├── SKILL.md
├── scripts/
│   ├── wps.py
│   └── wps_skill/
│       ├── __init__.py
│       ├── action_manifest.py
│       ├── addon_installer.py
│       ├── errors.py
│       ├── locking.py
│       ├── polling.py
│       ├── protocol.py
│       └── validation.py
├── references/
│   ├── action-manifest.json
│   ├── excel.md
│   ├── word.md
│   ├── powerpoint.md
│   ├── common.md
│   └── troubleshooting.md
└── assets/
    └── wps-addin/
        ├── manifest.xml
        ├── ribbon.xml
        ├── main.js
        └── handlers/
```

开发仓库可以在技能目录外保留 `tests/`、迁移脚本、ADR 和分析文档；最终分发目录只包含智能体完成 WPS 工作所需的资源。

## 4. 内部调用契约

`scripts/wps.py` 是智能体使用的内部入口，不是用户需要安装或管理的独立应用。

建议支持三个内部操作：

```text
check     检查 Python、WPS、加载项和端口状态
install   幂等安装或更新 WPS Add-in
invoke    执行一个 WPS Action
```

`invoke` 接收一个 JSON 请求：

```json
{
  "action": "getCellValue",
  "params": {
    "sheet": "Sheet1",
    "cell": "A1"
  },
  "timeout_ms": 30000,
  "confirmed": false
}
```

标准输出只产生一个 JSON 结果：

```json
{
  "ok": true,
  "action": "getCellValue",
  "data": {
    "value": 42
  }
}
```

失败结果使用稳定错误码：

```json
{
  "ok": false,
  "action": "getCellValue",
  "error": {
    "code": "ADDIN_NOT_READY",
    "message": "WPS Add-in 尚未连接",
    "retryable": true
  }
}
```

诊断日志写入标准错误。失败时进程返回非零退出码，但仍尽可能在标准输出返回结构化错误。

## 5. Action manifest

`references/action-manifest.json` 是能力契约的唯一事实来源。每项至少包含：

```json
{
  "action": "deleteSheet",
  "application": "excel",
  "description": "删除工作表",
  "parameters": {},
  "required": ["sheet"],
  "result": {},
  "risk": "destructive",
  "requires_active_document": true
}
```

manifest 用于：

- Python Action 白名单；
- 纯标准库参数校验；
- Skill references；
- Action 数量生成；
- JS dispatch 覆盖检查；
- 风险确认；
- 行为测试参数化。

迁移期间建立一次性的“旧 MCP Tool → WPS Action”对照表，用来证明能力没有遗漏；该对照表不进入最终接口。

## 6. Python runner

每次 `invoke`：

1. 定位技能根目录；
2. 加载并校验 action manifest；
3. 检查 Action 是否存在；
4. 校验必填参数和基础类型；
5. 对 `destructive` Action 校验 `confirmed=true`；
6. 获取跨平台文件锁；
7. 检查加载项版本和注册状态；
8. 在 `127.0.0.1:58891` 临时启动标准库 HTTP Server；
9. 等待 WPS Add-in 发起 `GET /poll`；
10. 返回带唯一 `requestId` 的 Action；
11. 接收 `POST /result`；
12. 校验 `requestId` 并输出 JSON；
13. 关闭服务器、释放文件锁并退出。

需要显式处理：

- 端口占用；
- 并发调用；
- 加载项未安装；
- 加载项已安装但 WPS 未重启；
- WPS 未启动；
- 轮询超时；
- 非法 JSON；
- 结果属于其他请求；
- Python 进程异常退出后的锁恢复。

## 7. 加载项安装

`addon_installer.py` 在首次使用或版本变化时运行。

目标路径：

- Linux：优先检测 `~/.local/share/Kingsoft/wps/jsaddons/`，并兼容实际 WPS 安装使用的用户级路径；
- Windows：`%APPDATA%\kingsoft\wps\jsaddons\`。

安装要求：

- 使用统一、稳定且以 `_` 结尾的加载项目录名；
- 复制 `assets/wps-addin/`，不依赖源码仓库其他路径；
- 通过版本或摘要判断是否需要更新；
- 原子更新，失败时保留旧版本；
- 备份并合并 `publish.xml`；
- 不覆盖其他 `<jsplugin>`；
- `publish.xml` 非法时停止并给出可操作错误，不静默重建；
- 安装或更新后返回 `restart_required=true`；
- 不请求管理员权限。

回环服务应使用每次安装生成的本地共享凭证，避免任意网页读取待执行 Action。凭证保存在用户级配置中并注入已安装加载项；Python 仅监听 `127.0.0.1`。

## 8. 统一 WPS Add-in

统一加载项以 `wps-claude-assistant` 的轮询实现为基础，不使用当前仅显示 COM 状态的 `wps-claude-addon`。

当前静态对比结果：

```text
JavaScript Add-in dispatch：227
PowerShell dispatch：241
JavaScript 相对缺口：14
```

需要补齐的 dispatch：

```text
closeDocument
convertFormat
createDocument
enableTrackChanges
findInDocument
getDocumentParagraphs
getExcelContext
getTrackChangesStatus
insertSlidesFromFile
openFile
replaceBookmarkContent
replacePptImage
replaceRange
smartFillField
```

补齐时应优先移植行为和错误语义，不复制 PowerShell 结构。完成后由 manifest 与 dispatch 的集合测试确认无缺口。

## 9. Skill 内容迁移

唯一 `SKILL.md` 只保留：

- 任务识别和 Excel/Word/PowerPoint/跨应用路由；
- 首次使用 readiness 流程；
- 如何按需读取 references；
- Python runner 调用规范；
- 写操作和破坏性操作规则；
- 保存、覆盖、只读、超时和 WPS 重启处理；
- 多步骤任务逐个 Action 串行执行的规则。

详细能力、参数和示例分别放进应用 reference。不要在 `SKILL.md` 重复完整 Action 列表，也不要保留“调用 MCP 工具”等措辞。

## 10. 风险规则

manifest 风险分级：

- `read`：只读，无需确认；
- `write`：普通修改，用户请求已明确授权时执行；
- `destructive`：删除、覆盖、丢弃未保存工作等，必须显式确认。

Skill 负责对话确认，runner 负责强制执行。`destructive` Action 缺少 `confirmed=true` 时必须在接触 WPS 前失败。

## 11. 实施阶段

### 阶段 A：冻结行为基线

1. 提取所有 ToolDefinition、handler 调用的 Action 及参数转换；
2. 提取 JavaScript 和 PowerShell dispatch；
3. 解决当前工具数量和名称漂移；
4. 生成迁移对照表和第一版 manifest；
5. 为代表性现有行为保存 JSON fixture。

### 阶段 B：建立新 Skill 骨架

1. 使用 skill 初始化工具创建单一 `wps-office`；
2. 建立 `scripts/`、`references/` 和 `assets/`；
3. 编写精简 `SKILL.md`；
4. 验证技能元数据和目录结构。

### 阶段 C：实现 Python 通信

1. 实现错误模型和 JSON 输入输出；
2. 实现 manifest 加载与校验；
3. 实现文件锁、临时 HTTP Server 和超时；
4. 实现 `check`、`install`、`invoke`；
5. 用模拟 Add-in 完成协议测试。

### 阶段 D：统一加载项

1. 整理轮询版加载项为平台无关资产；
2. 增加本地共享凭证；
3. 补齐 14 个 Action；
4. 修正旧 MCP、Claude、macOS 专用命名；
5. 让 manifest 与 dispatch 集合完全一致。

### 阶段 E：迁移 Skill 知识

1. 合并现有四个 Skill；
2. 删除 MCP Tool 术语和调用示例；
3. 按应用生成 references；
4. 加入 readiness、安全和错误恢复流程；
5. 用真实用户任务进行前向测试。

### 阶段 F：跨平台验收

在四个目标分别安装对应架构 WPS，并执行同一套 smoke tests：

- 首次安装加载项；
- 重复安装不改变有效配置；
- WPS 重启提示和连接检查；
- Excel 新建、读、写、公式和保存；
- Word 新建、插入、读取和保存；
- PowerPoint 新建、添加幻灯片、写文本和保存；
- Excel 数据到 PowerPoint 的跨应用流程；
- 破坏性 Action 无确认时被拒绝；
- 超时、端口占用和并发锁；
- 14 个补齐 Action 的针对性测试。

### 阶段 G：删除旧运行时

只有在 manifest、自动测试和四平台 smoke tests 全部通过后，才删除：

- `wps-office-mcp/`；
- 根 `package.json` 中的 Node/MCP 内容；
- MCP/Node 安装和开发脚本；
- Windows COM-only 加载项；
- PowerShell COM 脚本；
- 三个独立应用 Skill；
- README、INSTALL 和注释中的 MCP 配置说明。

删除前保留可回溯的迁移对照和 Git 历史，不采用一次性大爆炸替换。

## 12. 验收标准

- 最终 `wps-office` 目录不包含 Node.js、npm、TypeScript、MCP SDK 或 MCP 配置依赖；
- 目标智能体只需加载一个 `SKILL.md`；
- 首次调用可在用户权限下完成 Add-in 安装；
- 同一 Python 源码运行于四个平台；
- manifest 中每个 Action 都有且只有一个 Add-in dispatch；
- 旧能力到 Action 的迁移对照无遗漏；
- `read`、`write`、`destructive` 规则由 runner 测试覆盖；
- stdout、stderr、退出码和错误 JSON 契约稳定；
- 四个真实 WPS 环境的 smoke tests 全部通过；
- 项目文档不再把 Skill、Action、Add-in 和 MCP 混为一谈。
