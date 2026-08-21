# WPS Office Skill 开发者指南

本文面向需要理解、修改或扩展本仓库的开发者。它回答三个问题：项目如何运行、每类实现放在哪里、改动后如何确认没有漏改。

## 先建立整体认识

本仓库交付的不是常驻服务，而是一个可独立复制的 [WPS Skill Package](skills/wps-office/)。Agent Host 加载 [`SKILL.md`](skills/wps-office/SKILL.md)，按需启动 Python Runner；Runner 再通过本机回环 HTTP 与运行在 WPS Office 内的 JavaScript Add-in 通信。一次 Python 进程只执行一个 WPS Action，完成后退出。

项目当前公开 249 个 WPS Actions。数量、参数、结果、应用归属、前置条件和风险等级的唯一事实来源是 [`action-manifest.json`](skills/wps-office/references/action-manifest.json)，而不是 README、旧 MCP 名称或 JavaScript 函数名。迁移时的数量和历史映射见 [`action-baseline.md`](doc/action-baseline.md)；领域术语见 [`CONTEXT.md`](CONTEXT.md)。

```text
Agent Host
  │ 加载 SKILL.md；执行 check / invoke
  ▼
Python Runner (wps.py)
  ├─ 安装或更新 Add-in
  ├─ 读取 Action manifest，校验参数和风险
  ├─ 获取跨进程锁
  └─ 在 127.0.0.1:58891 临时监听
             ▲ /poll       │ /ack、/result
             │             ▼
WPS Add-in (main.js，HTTP 客户端)
  ├─ 使用本机凭证认证
  ├─ 按 requestId 确认和关联请求
  ├─ dispatch WPS Action
  └─ 调用 WPS JavaScript 对象模型
             │
             ▼
Excel / Word / PowerPoint
```

这条链路的架构出处是：Skill Package 交付形态见 [ADR-0002](docs/adr/0002-distribute-as-an-agent-skill-package.md)，统一 JavaScript Platform Bridge 见 [ADR-0005](docs/adr/0005-use-one-wps-js-bridge-on-all-targets.md)，单 Action 单进程见 [ADR-0008](docs/adr/0008-run-one-action-per-python-process.md)，Action 契约边界见 [ADR-0010](docs/adr/0010-use-wps-action-identifiers-as-the-capability-contract.md)。

## 仓库地图

| 位置 | 职责 | 什么时候改 |
|---|---|---|
| [`skills/wps-office/SKILL.md`](skills/wps-office/SKILL.md) | Agent 的入口说明：就绪检查、Action 选择、调用、安全和错误恢复 | Agent 的操作流程或公开行为变化时 |
| [`references/action-manifest.json`](skills/wps-office/references/action-manifest.json) | 所有 WPS Action 的规范契约 | Action 增、删、改时 |
| [`references/common.md`](skills/wps-office/references/common.md) | 跨应用文件工作流 | 通用 Action 或推荐工作流变化时 |
| [`references/excel.md`](skills/wps-office/references/excel.md) | Excel 渐进式工作流和 Action 目录 | Excel Action 或工作流变化时 |
| [`references/word.md`](skills/wps-office/references/word.md) | Word 渐进式工作流和 Action 目录 | Word Action 或工作流变化时 |
| [`references/powerpoint.md`](skills/wps-office/references/powerpoint.md) | PowerPoint 渐进式工作流和 Action 目录 | PowerPoint Action 或工作流变化时 |
| [`scripts/wps.py`](skills/wps-office/scripts/wps.py) | Runner 进程边界、契约校验、风险门禁、锁和回环协议 | 通用调用机制变化时；普通 Action 通常不用改 |
| [`scripts/wps_skill/addon_installer.py`](skills/wps-office/scripts/wps_skill/addon_installer.py) | 平台识别、用户级 Add-in 安装、注册、凭证、摘要和事务恢复 | 安装位置、平台支持或安装流程变化时 |
| [`assets/wps-addin/main.js`](skills/wps-office/assets/wps-addin/main.js) | Action dispatch 与所有真实 WPS 行为 | Action 实现增、删、改时 |
| [`assets/wps-addin/index.html`](skills/wps-office/assets/wps-addin/index.html) | Add-in 浏览器入口和脚本加载顺序 | 启动资源变化时 |
| [`assets/wps-addin/manifest.xml`](skills/wps-office/assets/wps-addin/manifest.xml) | WPS Add-in 元数据、宿主和资源声明 | Add-in 注册信息变化时 |
| [`assets/wps-addin/ribbon.xml`](skills/wps-office/assets/wps-addin/ribbon.xml) | WPS 功能区 UI | 按钮或 Ribbon 行为变化时 |
| [`scripts/validate_action_manifest.py`](scripts/validate_action_manifest.py) | manifest、dispatch 和冻结迁移证据的静态一致性校验 | 契约 Schema 或校验规则变化时 |
| [`tests/`](tests/) | Runner、安装器、Action 契约和独立分发的黑盒测试 | 与对应实现一起改 |
| [`docs/adr/`](docs/adr/) | 难以逆转的系统级设计决定 | 架构边界或兼容策略发生实质变化时 |
| [`doc/migration/legacy-tool-action-map.json`](doc/migration/legacy-tool-action-map.json) | 已完成迁移的冻结审计证据 | 不作为日常 Action 配置修改 |

生产分发边界是整个 [`skills/wps-office/`](skills/wps-office/) 目录。根目录下的校验器、测试、迁移材料和 ADR 是维护仓库所需资源，不应成为分发包的运行时依赖；这个约束由 [`test_package_integrity.py`](tests/test_package_integrity.py) 检查。

## 一次调用如何完成

### 1. Agent 选择 Action

Agent 先执行 `python3 scripts/wps.py check`，然后读取 manifest 和相关应用参考文档，最后执行：

```bash
python3 scripts/wps.py invoke '{"action":"getCellValue","params":{"sheet":"Sheet1","row":1,"col":1},"timeout_ms":30000}'
```

公开调用方式和恢复规则来自 [`SKILL.md`](skills/wps-office/SKILL.md)。

### 2. Runner 安装并检查 Add-in

[`addon_installer.py`](skills/wps-office/scripts/wps_skill/addon_installer.py) 检测操作系统和 CPU 架构，在当前用户的 WPS 配置目录安装 Add-in，并合并而不是覆盖 `publish.xml`。它还生成不进入源码目录的 `wps-skill-config.js`，其中包含本机凭证和安装摘要。

Add-in 资源发生变化后，资源摘要会变化。Runner 会更新已安装副本并返回 `restart_required`，直到 WPS 重启后加载的新 Add-in 用相同摘要完成认证 ping。安装过程、损坏修复、原子升级和平台路径由 [`test_addon_installer.py`](tests/test_addon_installer.py) 覆盖。

### 3. Runner 在 WPS 之前拦截无效请求

[`wps.py`](skills/wps-office/scripts/wps.py) 依次完成：

1. 从 manifest 查找 Action；未知名称返回 `UNKNOWN_ACTION`。
2. 根据 manifest 校验 `params`。当前支持 `type`、`enum`、`minimum`、`maximum`、`pattern`、对象必填字段、`additionalProperties`、数组元素和 `anyOf`。
3. 对打开、另存和 PDF 输出等已知文件 Action 做本地路径规范化及存在性检查。
4. 应用 `read`、`write`、`destructive` 风险门禁；`destructive` 没有 `confirmed: true` 时不会送达 WPS。规则出处是 [ADR-0011](docs/adr/0011-require-confirmation-for-destructive-actions.md)。
5. 获取当前用户的非阻塞跨进程锁。`check` 和 `invoke` 共用该锁，竞争者收到 `ACTION_BUSY`。

manifest 中的 `prerequisites` 是 Action 选择和契约元数据。Runner 会统一检查 WPS 与 Add-in 是否就绪，但不会逐项解释 `active_workbook` 等应用状态；具体活动对象仍由 Add-in handler 检查并返回 WPS 错误。

### 4. Runner 与 Add-in 交换一次请求

Runner 只绑定 `127.0.0.1:58891`，Add-in 是主动轮询的 HTTP 客户端：

- `GET /poll` 获取候选 Action；
- `POST /ack` 表示当前 `requestId` 已被接收；
- `POST /result` 上传同一 `requestId` 的执行结果；
- 三条受保护路由都使用安装时生成的 Bearer 凭证；
- CORS `OPTIONS` 只协商请求，不读取或改变 Action 状态。

请求关联和错误分类实现在 [`wps.py`](skills/wps-office/scripts/wps.py)，浏览器侧协议实现在 [`main.js`](skills/wps-office/assets/wps-addin/main.js)。对应黑盒测试位于 [`test_wps_runner.py`](tests/test_wps_runner.py)。Action 已确认送达后，写入或破坏性操作超时不可自动重试，因为 WPS 可能在 Runner 超时后完成原操作。

### 5. Add-in dispatch 并返回结果

[`main.js`](skills/wps-office/assets/wps-addin/main.js) 中的 `handleAction` 用 `switch` 将 Action 名称分发给具体 `handle...` 函数。新 handler 应统一返回：

```javascript
// 成功
return { success: true, data: { /* 必须匹配 manifest 的 result */ } };

// 失败
return { success: false, error: "可诊断的 WPS 错误" };
```

Runner 会再次使用 manifest 校验成功结果中的 `data`，不匹配时返回 `INVALID_RESULT`；WPS 侧失败转换为 `WPS_ACTION_FAILED`。因此 manifest 不只是说明文档，同时约束进入 WPS 的参数和离开 WPS 的结果。

## 如何增加一个 WPS Action

普通 Action 功能的完整改动链是：

```text
manifest 契约
  → reference group / 应用参考文档
  → main.js dispatch
  → main.js handler
  → 契约测试与 Runner 黑盒测试
```

### 1. 先定义规范契约

在 [`action-manifest.json`](skills/wps-office/references/action-manifest.json) 的 `actions` 中添加 lowerCamelCase 名称，并加入一个 `reference_groups` 分组。最小结构如下：

```json
{
  "action": "getSomething",
  "application": "excel",
  "description": "Read something from WPS Excel.",
  "parameters": {
    "type": "object",
    "properties": {},
    "required": [],
    "additionalProperties": false
  },
  "result": {
    "type": "object",
    "properties": {
      "value": { "type": "string" }
    },
    "required": ["value"],
    "additionalProperties": false,
    "description": "Result returned by the getSomething WPS Action."
  },
  "prerequisites": ["wps_running", "active_workbook"],
  "risk": "read"
}
```

选择 `risk` 时按实际后果分类：读取用 `read`；普通、可预期的编辑用 `write`；删除、覆盖、清空或可能丢失内容的操作用 `destructive`。不要把确认逻辑写进单个 handler，统一门禁在 Runner。

参数必须描述 handler 真正无条件读取的字段，结果类型必须与 WPS 实际返回值一致。复杂对象需要递归声明 `properties`；数组需要声明 `items`。如果需要新的 Schema 关键字，必须同时扩展 Runner 和仓库校验器，不能只把关键字写进 JSON。

### 2. 实现 WPS 行为

在 [`main.js`](skills/wps-office/assets/wps-addin/main.js) 的 `handleAction` 中增加一个 `case`，再添加对应 handler。尽量复用同一应用附近已有的对象查找、索引、颜色、路径和结果转换方式。

WPS 的集合索引通常是一基，但 Word 字符范围等契约可能是零基且右端不包含。不要从相邻 Action 猜测；以对应参考文档、manifest 和现有 handler 为准，并用测试锁定边界。

### 3. 更新给 Agent 的渐进式文档

把 Action 加入相应参考文件中标记包围的目录：

- common：[`common.md`](skills/wps-office/references/common.md)
- Excel：[`excel.md`](skills/wps-office/references/excel.md)
- Word：[`word.md`](skills/wps-office/references/word.md)
- PowerPoint：[`powerpoint.md`](skills/wps-office/references/powerpoint.md)

若 Action 改变推荐工作流，还要更新标记目录之外的步骤说明。精确参数和结果仍只维护在 manifest，不要在 Markdown 中复制一份容易失真的完整 Schema。

### 4. 添加最接近风险的测试

- 契约、分组、文档目录或 dispatch 一致性：[`test_action_manifest.py`](tests/test_action_manifest.py)
- 参数在接触 Add-in 前被拒绝、风险门禁、结果校验、通信行为：[`test_wps_runner.py`](tests/test_wps_runner.py)
- 独立复制后的资源解析或协议约束：[`test_package_integrity.py`](tests/test_package_integrity.py)
- 安装位置、升级、注册或摘要行为：[`test_addon_installer.py`](tests/test_addon_installer.py)

[`representative-actions.json`](tests/fixtures/representative-actions.json) 只保存跨应用和风险类型的代表样例，不要求每个新 Action 都添加一项。Action 自身需要契约样例时，可在 manifest 项中加入 `examples`。

## 如何修改或删除 WPS Action

### 修改

修改参数、结果或风险等级时，同时检查以下位置：

1. manifest 中的 Action 契约和 `examples`；
2. `main.js` 的 handler；
3. 应用参考文档中的工作流描述；
4. 针对该行为的契约测试和 Runner 黑盒测试；
5. 若属于文件输入或输出，检查 `wps.py` 的 `prepare_file_action_params` 是否也要支持。

只改 handler 而不改 manifest，常见结果是 Runner 拒绝新参数或把新结果判为 `INVALID_RESULT`；只改 manifest 而不改 handler，则静态校验可能通过，但真实 WPS 行为和契约会分叉。

### 删除

WPS Action 是公开能力标识，删除不是普通的死代码清理。按本仓库的“项目内兼容”约定，先用 GitHub Issue 和 ADR 记录影响、替代方案及迁移步骤；领域名称是 [`WPS Action Retirement`](CONTEXT.md)。

实施删除时至少同步处理：

1. manifest 的 Action 项和所属 `reference_groups`；
2. `main.js` 的 dispatch `case`，以及不再被复用的 handler；
3. 应用参考文档的目录和工作流；
4. 对应的固定契约、样例和黑盒测试。

[`legacy-tool-action-map.json`](doc/migration/legacy-tool-action-map.json) 是历史迁移证据，不是当前 Action 开关。当前 [`validate_action_manifest.py`](scripts/validate_action_manifest.py) 仍要求历史映射指向现存 Action，并把 ADR-0014 的退役集合固定在代码中。因此，若要删除被旧映射引用的规范 Action，需要在同一 Issue/ADR 中先调整“冻结历史证据如何与当前 Action 集合解耦”的校验方式；不要直接篡改历史映射来让测试变绿。

## 其他功能应该在哪里改

| 想改的功能 | 主要实现 | 首要测试 | 容易漏掉的同步项 |
|---|---|---|---|
| Runner 请求格式、超时、错误码 | [`wps.py`](skills/wps-office/scripts/wps.py) | [`test_wps_runner.py`](tests/test_wps_runner.py) | `SKILL.md`、包内 README |
| `/poll`、`/ack`、`/result` 或认证 | [`wps.py`](skills/wps-office/scripts/wps.py)、[`main.js`](skills/wps-office/assets/wps-addin/main.js) | `test_wps_runner.py`、`test_package_integrity.py` | 请求头、requestId、CORS、敏感信息策略 |
| Add-in 安装目录或平台识别 | [`addon_installer.py`](skills/wps-office/scripts/wps_skill/addon_installer.py) | [`test_addon_installer.py`](tests/test_addon_installer.py) | `INSTALL.md`、包内 README、平台 ADR |
| Add-in 升级和注册事务 | [`addon_installer.py`](skills/wps-office/scripts/wps_skill/addon_installer.py) | `test_addon_installer.py` | `publish.xml` 合并、备份恢复、权限和摘要 |
| Ribbon 或 Add-in 元数据 | [`ribbon.xml`](skills/wps-office/assets/wps-addin/ribbon.xml)、[`manifest.xml`](skills/wps-office/assets/wps-addin/manifest.xml) | `test_package_integrity.py` | [`assets/wps-addin/README.md`](skills/wps-office/assets/wps-addin/README.md) |
| Agent 使用流程 | [`SKILL.md`](skills/wps-office/SKILL.md) | `test_wps_runner.py`、`test_package_integrity.py` | 根 README、INSTALL、应用参考文档 |
| manifest Schema 或校验规则 | [`wps.py`](skills/wps-office/scripts/wps.py)、[`validate_action_manifest.py`](scripts/validate_action_manifest.py) | `test_action_manifest.py`、`test_wps_runner.py` | 两侧必须接受和拒绝相同契约 |
| 分发包结构 | [`skills/wps-office/`](skills/wps-office/) | `test_package_integrity.py`、`test_action_manifest.py` | `.claude-plugin/marketplace.json`、安装文档 |
| 真实环境准备 | [`setup-wps-validation-environments.sh`](scripts/setup-wps-validation-environments.sh) | 目标机器上的 smoke test | 不把凭证、SSH 目标或用户数据写入记录 |

修改 Python 运行时要遵守 [ADR-0007](docs/adr/0007-use-only-the-python-standard-library.md)：支持 Python 3.9+，只使用标准库。修改运行模型要遵守 [ADR-0008](docs/adr/0008-run-one-action-per-python-process.md)：不要为了复用状态引入常驻 daemon。修改平台声明时先读 [ADR-0004](docs/adr/0004-support-four-64-bit-platform-targets.md)、[ADR-0012](docs/adr/0012-require-real-wps-tests-on-each-target.md)、[ADR-0013](docs/adr/0013-recognize-macos-add-in-installation-as-experimental.md) 和 [ADR-0015](docs/adr/0015-separate-migration-completion-from-platform-certification.md)。

## 验证改动

在仓库根目录运行：

```bash
python3 scripts/validate_action_manifest.py
python3 -m unittest discover -s tests -v
```

第一条检查 manifest Schema、Action 唯一性、reference group、JavaScript dispatch 集合和冻结迁移映射；第二条覆盖契约、Runner 回环协议、风险门禁、安装器与独立分发。当前基线应显示 `249 WPS Actions`，完整测试为 170 项。

涉及真实 WPS 对象模型、安装目录或平台行为时，自动化测试只能证明仓库契约和模拟协议没有破坏。真实平台支持结论必须来自对应环境的 WPS smoke test，不能由静态测试或另一平台的成功推断。

## 常见失败如何定位

| 现象 | 通常原因 | 从哪里查 |
|---|---|---|
| `unmanifested WPS Action` | `main.js` 有 `case`，manifest 没有 Action | manifest 与 `handleAction` |
| `missing from the packaged JavaScript Add-in` | manifest 有 Action，dispatch 没有 `case` | `main.js` 的 `switch` |
| 参考目录测试失败 | reference group 与 Markdown 标记区不一致 | manifest 分组和对应应用 `.md` |
| `INVALID_PARAMS` | 请求不满足 manifest，或用了未声明字段 | manifest 的 `parameters` |
| `INVALID_RESULT` | handler 返回值与 manifest `result` 不一致 | handler 返回的 `data` 和结果 Schema |
| `CONFIRMATION_REQUIRED` | 破坏性 Action 未带显式确认 | manifest 的 `risk` 与调用请求 |
| `WPS_RESTART_REQUIRED` | Add-in 首次安装或资源摘要已变化 | 安装器、WPS 是否完全退出重启 |
| `ADDIN_NOT_READY` | 没有已认证 Add-in 轮询 | WPS 进程、Add-in 注册、凭证和加载顺序 |
| `ACTION_TIMEOUT` | Add-in 已确认，但 WPS 未在期限内完成 | 对应 handler；写操作不要自动重试 |
| `ACTION_BUSY` | 另一个 `check` 或 `invoke` 持有用户锁 | 等待当前进程退出，不要绕过锁 |
| `REQUEST_ID_MISMATCH` | 固定端口复用时收到旧进程请求 | `/ack`、`/result` 的 requestId 关联 |

## 提交前快速检查

- 改动是否位于正确的事实来源，而不是只改了 README？
- Action 的 manifest、reference group、dispatch、handler、参考文档和测试是否同步？
- handler 的成功 `data` 是否精确匹配结果 Schema？
- 删除、覆盖或丢弃内容的行为是否标为 `destructive`？
- 是否保持 Python 3.9+ 标准库、单 Action 单进程和仅回环通信？
- 是否避免在 stdout、诊断或测试记录中暴露凭证和 Action 参数？
- 两条仓库验证命令是否都通过？

仓库的 Issue 和规格使用 GitHub Issues；操作约定见 [`docs/agents/issue-tracker.md`](docs/agents/issue-tracker.md)。新增领域术语或架构决策时，分别更新 [`CONTEXT.md`](CONTEXT.md) 和 [`docs/adr/`](docs/adr/)。
