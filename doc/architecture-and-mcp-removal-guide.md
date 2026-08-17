# WPS Skills 项目架构与“去 MCP”任务导读

> 适合第一次接触本仓库的开发者。本文基于当前代码整理，重点说明项目组成、调用链、核心模块、阅读顺序，以及“改成不依赖 MCP，并支持 Linux、Windows x86 和 ARM”这项任务可能代表什么。

> 后续设计访谈已经确定 Python Skill 的目标边界、四平台矩阵和迁移步骤。实施时以 [`python-wps-skill-migration-plan.md`](./python-wps-skill-migration-plan.md) 及 [`docs/adr/`](../docs/adr/) 中的已接受决定为准；本文保留为旧架构分析和迁移背景。

## 1. 一句话认识项目

这个项目让 AI 根据 Skills 中的操作流程，调用一组 WPS 工具，再通过 Windows COM 或 Linux/macOS WPS 加载项真正操作 WPS Office。

当前实现可以拆成三层：

1. **Skills 层**：告诉 LLM 应该选什么工具、按什么步骤完成任务。
2. **工具执行层**：注册、查找和执行 Excel、Word、PPT、通用工具。
3. **WPS 平台桥接层**：把工具动作发送给不同平台上的 WPS。

MCP 只应该是“AI 如何调用工具执行层”的一种上游协议，但当前代码把 MCP 启动、工具注册和部分运行状态写在了同一个模块里，因此去掉 MCP 不能只删除一个 npm 依赖。

## 2. 目录结构与职责

```text
wps-skills/
├── skills/                       # 给 LLM 阅读的操作说明
│   ├── wps-office/               # 跨应用与通用场景
│   ├── wps-excel/                # Excel 场景
│   ├── wps-word/                 # Word 场景
│   └── wps-ppt/                  # PPT 场景
│
├── wps-office-mcp/               # 当前工具执行进程；不只是 MCP 代码
│   ├── src/index.ts              # 当前进程入口
│   ├── src/server/
│   │   ├── mcp-server.ts         # MCP、工具装配、部分缓存与生命周期
│   │   └── tool-registry.ts      # 工具注册与调用分发
│   ├── src/tools/                # Excel、Word、PPT、通用工具定义与处理器
│   ├── src/client/
│   │   ├── wps-client.ts         # 统一 WPS 调用入口与平台选择
│   │   ├── mac-poll-server.ts    # Linux/macOS HTTP 反向轮询桥
│   │   └── wps-keepalive.ts      # WPS 保活相关逻辑
│   ├── src/types/                # 工具与 WPS 请求/响应类型
│   ├── src/utils/                # 日志和错误类型
│   └── scripts/wps-com.ps1       # Windows PowerShell COM 实现
│
├── wps-claude-assistant/         # Linux/macOS WPS JS 加载项
│   ├── main.js                   # 轮询、命令路由
│   └── handlers/                 # Excel、Word、PPT、通用动作实现
│
├── wps-claude-addon/             # Windows WPS 加载项和界面声明
├── scripts/                      # 各平台安装、构建和开发脚本
├── INSTALL.md                    # 安装说明
└── README.md                     # 项目总览
```

需要特别注意：`wps-office-mcp` 这个目录名容易让人误以为里面只有 MCP Server。实际上，工具注册、工具处理器、WPS 客户端以及平台桥接都在这里。未来去掉 MCP 后，这部分大概率仍然需要保留，只是应当改名或重新分层。

## 3. 当前整体架构

### 3.1 决策和执行是两件事

Skills 不负责执行代码，它们是提供给 LLM 的自然语言说明。真正的调用链如下：

```text
用户提出自然语言需求
        ↓
LLM 读取相关 SKILL.md
        ↓
LLM 根据 Skill 选择工具名和参数
        ↓
AI 宿主的 MCP Client
        ↓ stdio / MCP（tools/list、tools/call）
WpsMcpServer
        ↓
ToolRegistry
        ↓
具体 ToolHandler
        ↓
WpsClient.invokeAction(action, params)
        ↓
平台实现真正操作 WPS
        ↓
结果沿原路径返回给 LLM
```

因此：

- **Skill 回答“应该怎么做”。**
- **工具处理器回答“这个能力如何映射成 WPS action”。**
- **平台桥接回答“当前操作系统上怎么把 action 交给 WPS”。**
- **MCP 回答“AI 宿主如何发现并调用这些工具”。**

如果去掉 MCP，Skills 不会自动获得执行能力，必须提供 CLI、本地函数接口、HTTP 或宿主专用工具接口中的至少一种替代入口。

### 3.2 MCP 前后衔接的内容

MCP 上游是支持 MCP 的 AI 宿主，例如 Claude Code、Cursor 或 Codex。MCP Server 负责：

- 初始化和能力声明；
- 返回工具清单；
- 接收工具名及参数；
- 把成功结果或错误转换成 MCP 响应。

MCP 下游不是 WPS 本身，而是仓库内部的 `ToolRegistry`。随后才依次进入工具处理器、`WpsClient` 和平台桥接。

当前 SDK 只在 [`mcp-server.ts`](../wps-office-mcp/src/server/mcp-server.ts) 中直接导入：

- `Server`
- `StdioServerTransport`
- `ListToolsRequestSchema`
- `CallToolRequestSchema`
- `ErrorCode`
- `McpError`

这说明 `@modelcontextprotocol/sdk` 的直接引用面很小，但 `mcp-server.ts` 同时承担了额外职责，所以删除 SDK 后还要重新安置这些职责。

### 3.3 各平台执行链

#### Windows

```text
ToolHandler
  → WpsClient
  → spawn("powershell")
  → wps-com.ps1
  → Ket/Kwps/Kwpp COM 对象
  → WPS Office
```

Windows 的核心实现位于 [`wps-com.ps1`](../wps-office-mcp/scripts/wps-com.ps1)。它通过 COM 获取或创建 WPS 表格、文字和演示对象。

#### Linux/macOS

```text
ToolHandler
  → WpsClient
  → MacPollServer（名称仍带 Mac，但 Linux 也使用）
  → 本机 127.0.0.1:58891
  ← WPS JS 加载项每 500ms 轮询
  → main.js/handlers 执行 WPS JS API
  → POST /result 返回结果
```

对应实现是 [`mac-poll-server.ts`](../wps-office-mcp/src/client/mac-poll-server.ts) 和 [`wps-claude-assistant/main.js`](../wps-claude-assistant/main.js)。这里的 HTTP 轮询不是 MCP；即使删除 MCP，为了在 Linux 上与 WPS 加载项通信，这条桥仍可能需要保留。

## 4. 核心模块与关键抽象

| 模块 | 当前职责 | 是否依赖 MCP |
|---|---|---|
| `skills/*/SKILL.md` | LLM 的场景识别、步骤和工具使用说明 | 文案中硬编码了 MCP 工具概念 |
| `src/index.ts` | 创建并启动 MCP Server、处理进程退出 | 是 |
| `mcp-server.ts` | MCP 协议、工具装配、内置工具、跨应用缓存 | 是，而且职责过多 |
| `tool-registry.ts` | 注册、列出和调用工具 | 业务上不需要 MCP，但命名/错误类型带 MCP 痕迹 |
| `src/tools/**` | 工具定义和 ToolHandler | 基本不依赖 MCP SDK |
| `wps-client.ts` | 统一 action 调用并选择平台路径 | 否 |
| `mac-poll-server.ts` | Linux/macOS 与加载项通信 | 否 |
| `wps-com.ps1` | Windows COM 自动化 | 否 |
| `wps-claude-assistant/**` | Linux/macOS WPS 内部执行 | 否 |

### 4.1 ToolDefinition / ToolHandler

一个工具主要由两部分组成：

- `ToolDefinition`：名称、说明、输入 Schema、分类；
- `ToolHandler`：接收参数、调用 WPS、返回统一结果。

这些类型在 [`src/types/tools.ts`](../wps-office-mcp/src/types/tools.ts)，工具调度在 [`tool-registry.ts`](../wps-office-mcp/src/server/tool-registry.ts)。这是最适合保留并进一步稳定的核心抽象。

### 4.2 Tool 名称和 WPS action

LLM 看到的是类似 `wps_excel_get_cell_value` 的工具名；平台桥看到的通常是类似 `getCellValue` 的 action。工具处理器负责两者之间的映射。

理解一个完整调用时，不要只看工具定义，要沿下面的链追踪：

```text
工具名 → ToolHandler → WpsClient.invokeAction → action → 平台 action 实现
```

### 4.3 WpsClient

[`wps-client.ts`](../wps-office-mcp/src/client/wps-client.ts) 是当前最重要的平台抽象：

- Windows 进入 PowerShell COM；
- Linux/macOS 进入 HTTP 反向轮询。

不过平台判断当前是在模块加载时通过 `os.platform()` 完成，且没有读取 `process.arch`。所以它只做了操作系统分支，没有形成 CPU 架构支持矩阵。

### 4.4 Skills

Skills 是 LLM 的操作策略，不是运行时驱动。以 Excel Skill 为例，它要求 LLM 先获取工作表上下文，再调用相应工具完成任务。当前 Skill 中大量使用“调用 MCP 工具”的表述；如果替换成 CLI，工具名称、参数格式、结果格式和调用示例都必须同步更新。

## 5. 模块依赖关系

```text
skills
  └─ 逻辑上依赖工具名称、参数 Schema 和可用能力

src/index.ts
  └─ mcp-server.ts
       ├─ @modelcontextprotocol/sdk
       ├─ tool-registry.ts
       ├─ src/tools/index.ts
       ├─ wps-client.ts
       └─ 跨应用内存缓存

src/tools/index.ts
  ├─ excel/**
  ├─ word/**
  ├─ ppt/**
  └─ common/**
       └─ 大多数最终调用 wps-client.ts

wps-client.ts
  ├─ Windows → scripts/wps-com.ps1
  └─ Linux/macOS → mac-poll-server.ts
                        └─ wps-claude-assistant/main.js + handlers/**
```

最重要的依赖方向是：工具执行层依赖 WPS 平台桥接；MCP 适配层依赖工具执行层。理想情况下，工具执行层不应该反过来依赖 MCP。

## 6. 建议从哪里开始阅读

不要一开始通读两百多个工具，先沿一个简单工具走通整条链。

### 第一阶段：先理解产品和 LLM 行为

1. [`README.md`](../README.md)：了解产品定位和安装方式。
2. [`skills/wps-excel/SKILL.md`](../skills/wps-excel/SKILL.md)：理解 LLM 如何选择、组合工具。
3. [`skills/wps-office/SKILL.md`](../skills/wps-office/SKILL.md)：理解跨应用协调。

注意：README、Skills 和工具汇总文件中的工具数量目前互相不一致，因此把它们当作架构说明，不要把文档里的数量当作最终事实。

### 第二阶段：走通工具执行主链

4. [`src/types/tools.ts`](../wps-office-mcp/src/types/tools.ts)：先看工具定义、请求和结果类型。
5. [`tool-registry.ts`](../wps-office-mcp/src/server/tool-registry.ts)：看工具如何注册和调用。
6. [`src/tools/index.ts`](../wps-office-mcp/src/tools/index.ts)：看各业务工具如何汇总。
7. 选择一个简单工具，例如 [`src/tools/excel/workbook.ts`](../wps-office-mcp/src/tools/excel/workbook.ts)，沿 handler 找到对应 action。

### 第三阶段：理解 MCP 在哪里

8. [`mcp-server.ts`](../wps-office-mcp/src/server/mcp-server.ts)：重点看构造函数、请求处理器以及 `start()`。
9. [`src/index.ts`](../wps-office-mcp/src/index.ts)：看进程启动和关闭。

读完这两处，应该能分清“工具执行能力”和“MCP 协议包装”。

### 第四阶段：理解平台差异

10. [`wps-client.ts`](../wps-office-mcp/src/client/wps-client.ts)：看操作系统分流。
11. Windows 路径：[`wps-com.ps1`](../wps-office-mcp/scripts/wps-com.ps1)。
12. Linux 路径：[`mac-poll-server.ts`](../wps-office-mcp/src/client/mac-poll-server.ts) → [`wps-claude-assistant/main.js`](../wps-claude-assistant/main.js) → [`handlers/`](../wps-claude-assistant/handlers/README.md)。
13. 最后再看 [`scripts/install.sh`](../scripts/install.sh)、[`scripts/install.ps1`](../scripts/install.ps1) 和 [`INSTALL.md`](../INSTALL.md)，理解部署假设。

## 7. 当前架构中值得注意的问题

### 7.1 MCP Adapter 和工具执行引擎耦合

`mcp-server.ts` 不仅处理 MCP，还负责：

- 注册全部工具；
- 定义 12 个内置工具；
- 保存跨应用缓存；
- 管理进程级服务生命周期。

这会让“替换 MCP”变成大范围改动。更合理的结构是独立出一个与传输协议无关的工具执行模块，再让 CLI、MCP 或其他 Adapter 调用它。

### 7.2 Skills 硬编码 MCP 术语

多个 `SKILL.md` 明确写着“调用 MCP 工具”。如果任务要求完全不依赖 MCP，代码、安装脚本、Skills 和文档需要一起迁移，仅删除依赖无法完成任务。

### 7.3 平台支持声明缺少架构维度

代码只判断 `darwin`、`linux` 和其他平台，没有明确处理：

- `ia32`（通常所说的 Windows x86/32 位）；
- `x64`；
- `arm`；
- `arm64`。

“TypeScript 能编译”不等于“WPS、Node、PowerShell COM 和加载项在该架构上通过验证”。

### 7.4 Linux/macOS 命名和实现混用

Linux 和 macOS 共用轮询实现，但代码仍使用 `MacPollServer`、`execMacPoll`、`MAC_POLL_PORT` 等名称。这会增加阅读和后续平台扩展成本。

### 7.5 轮询桥一次只能保存一个待执行命令

`mac-poll-server.ts` 使用单个 `pendingCommand`，新的并发调用可能覆盖尚未完成的旧调用。若新入口允许并发，必须明确串行化或实现队列。

### 7.6 文档和代码中的工具数量漂移

当前至少存在以下不同说法：

- 根 README：231 个专业工具，加 12 个内置工具；
- `src/tools/index.ts` 注释：235 个专业工具，加 12 个内置工具；
- Skills 中也有其他数量。

工具清单应由代码生成或通过测试校验，避免人工维护多个数字。

### 7.7 全局单例和模块级状态较多

`ToolRegistry`、`wpsClient`、轮询服务器、PPT 目标以及跨应用缓存都有全局或静态状态。这使并发调用、隔离测试和多实例运行更加困难。

## 8. 关于 `@modelcontextprotocol/sdk` 的现状

当前版本锁定为 `1.29.0`。本地测量结果如下：

| 内容 | 大小/数量 |
|---|---:|
| npm 压缩包 | 约 559 KiB |
| SDK 解压后 | 约 5.84 MiB、677 个文件 |
| SDK 加生产依赖 | 约 24.1 MiB、93 个包、3507 个文件 |
| 官方 TypeScript `src` 源码 | 约 1.16 MiB、90 个文件 |

仓库只直接使用该 SDK 的少量 server/stdio/schema 能力，但完整单体包还包含 Client、HTTP、SSE、OAuth、Schema 校验、experimental 功能，以及 ESM/CommonJS 两套构建。

可以把 MIT 许可的源码迁入仓库，但这相当于维护内部 fork，必须保留许可证和上游版本信息。更重要的是：**复制 SDK 源码仍然是在使用 MCP，并不满足“完全不依赖 MCP”的要求。**

## 9. 上一个问题及答案整理

### 问题

> 我现在是接到任务，让我把这个仓库改成不依赖 MCP 的，支持 Linux、Windows x86，以及 ARM。他是什么意思？

### 核心答案

这句话把两个维度放在了一起：

1. **“不依赖 MCP”是上游调用协议和集成方式的变化。**
2. **“支持 Linux、Windows x86、ARM”是操作系统、CPU 架构和交付方式的变化。**

任务方大概率希望把 WPS 工具能力从 MCP Server 中抽出来，提供一个不需要 MCP 客户端的新调用入口，并让交付产物能在指定平台运行。

可能的目标架构是：

```text
LLM
  ↓ 读取 Skills
本地 CLI / 宿主提供的工具接口
  ↓
与协议无关的 WPS Tool Engine
  ↓
平台 Adapter
  ├─ Windows → PowerShell COM → WPS
  └─ Linux   → HTTP 轮询加载项 → WPS
```

例如，新入口可能是：

```bash
wps-cli call wps_excel_read_range '{"sheet":"Sheet1","range":"A1:B10"}'
```

但是任务原话不足以直接开始实现，因为“不依赖 MCP”至少有三种不同含义：

| 含义 | 改动范围 |
|---|---|
| 只删除 `@modelcontextprotocol/sdk` | 自己实现 MCP 或复制源码，本质仍使用 MCP |
| 不使用 MCP 协议 | 删除 MCP Server，必须增加 CLI/HTTP/宿主接口 |
| Skills 独立运行 | 不成立；Skills 本身不会执行 WPS，仍需执行入口 |

平台表述也有歧义：

- “Windows x86”可能特指 32 位 `ia32`，也可能只是泛指 Intel/AMD Windows；
- “ARM”可能是 Linux ARM64、Linux ARMv7 或 Windows ARM64；
- 还需要确认是否同时要求 Linux x64、Windows x64；
- 需要确认用户是否必须预装 Node.js，还是要提供独立可执行文件。

### 应向任务方确认的问题

可以直接使用下面这段话：

> 我理解目标是：删除 MCP 协议和 `@modelcontextprotocol/sdk` 依赖，将工具注册与 WPS 执行能力提取成独立运行模块，再通过新的 CLI 或本地接口供 Skills 调用。请确认：  
> 1. “不依赖 MCP”是只去掉 npm SDK，还是连 MCP 协议、MCP Server 和 MCP 配置都删除？  
> 2. 去掉 MCP 后，Skills 通过 CLI、HTTP 还是其他宿主接口执行 WPS 操作？  
> 3. 目标平台具体是 Windows ia32、Windows x64、Windows ARM64、Linux x64、Linux ARM64 中的哪些？  
> 4. 是否要求免安装 Node.js 的独立可执行产物？  
> 5. 每个平台对应的 WPS 版本和安装包是什么？

## 10. 建议的任务边界和验收标准

需求澄清后，建议把任务拆成以下可验收部分。

### 10.1 与协议无关的工具执行模块

建议先形成一个小而稳定的接口，例如：

```ts
interface WpsToolEngine {
  listTools(): ToolDefinition[];
  callTool(name: string, args: Record<string, unknown>): Promise<ToolCallResult>;
}
```

它负责工具装配、调用和公共状态，但不读取 stdio、不理解 MCP，也不直接决定平台协议。

### 10.2 替代 MCP 的调用 Adapter

任务方必须选定至少一种：

- CLI：最容易被 Skills 和本地 Agent 调用；
- 本地 HTTP：适合常驻进程，但需要端口、鉴权和生命周期设计；
- JavaScript/TypeScript 库：适合同进程宿主；
- 特定 Agent 平台的原生工具插件：会绑定具体宿主。

### 10.3 平台 Adapter

至少明确并验证：

| 操作系统 | CPU 架构 | WPS 版本 | Node/运行时 | 执行方式 | 是否验收通过 |
|---|---|---|---|---|---|
| Windows | ia32/x64/arm64 待确认 | 待确认 | 待确认 | PowerShell COM | 待验证 |
| Linux | x64/arm/arm64 待确认 | 待确认 | 待确认 | WPS JS 加载项轮询 | 待验证 |

### 10.4 最低验收项

- 生产依赖和源码中不再引用 MCP（如果要求完全去 MCP）；
- 安装流程不再注册 MCP Server；
- Skills 使用新入口，能够列出并调用工具；
- 新入口具有稳定的 JSON 参数、结果和错误格式；
- 每个目标 OS/CPU 组合均能安装、启动并执行最小 smoke test；
- 至少验证 Excel 读写、Word 文本读写、PPT 新增幻灯片和通用保存；
- 并发策略明确，不能静默覆盖待执行命令；
- 工具列表由代码生成或测试校验；
- 文档明确运行时、WPS 版本和安装要求。

## 11. 当前最稳妥的下一步

在未获得任务方答复前，不建议直接删除 `mcp-server.ts`。先完成下面三件事：

1. 确认“不依赖 MCP”的准确含义和替代调用入口；
2. 确认完整的平台矩阵及是否需要独立可执行产物；
3. 把 `ToolRegistry + allTools + 内置工具/缓存` 提炼成与 MCP 无关的工具执行模块设计。

这样后续无论选择 CLI、HTTP 还是宿主插件，都可以复用同一套 WPS 工具实现，而不需要重写两百多个工具。
