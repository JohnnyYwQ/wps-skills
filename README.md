# WPS Skills

> 此项目的任何功能、架构更新，必须在结束后同步更新相关文档。这是我们契约的一部分。

AI编程助手通过自然语言操控WPS Office的MCP工具集。

## Python Skill Package 迁移状态

Issue #4 已在 `skills/wps-office/` 交付首个可加载的 WPS Skill Package 链路。Issue #5 进一步加入 Linux/Windows x86_64、ARM64 用户级 WPS Add-in 首次安装、幂等升级、`publish.xml` 安全合并、本机认证配置和就绪检查。Issue #6 加入按用户跨进程互斥、共享密钥认证、Action 全程超时、稳定传输错误和端口/锁清理。Issue #7 从 Action 清单强制执行 `read`、`write`、`destructive` 风险策略，并在 WPS Add-in 收到请求前拒绝未显式确认的破坏性操作。Issue #8 已迁移 Excel 核心数据工作流；Issue #9 继续覆盖格式、图表、数据分析、数据透视、保护、图片和上下文感知能力；Issue #10 已迁移 Word 文档生命周期、内容、范围、模板、格式、书签、批注和修订工作流；Issue #11 已迁移 PowerPoint 演示文稿、幻灯片、文本、图片、表格、备注和基础版式工作流，并补齐统一 Add-in 的页面尺寸能力；Issue #12 已迁移 PowerPoint 高级设计工作流，包括形状、图表、流程图、美化、动画、切换、母版、外部幻灯片和图片替换。各迁移切片均包含渐进式 reference、契约覆盖及 Runner/Fake Add-in 黑盒场景。智能体先运行 `python3 scripts/wps.py check`，仅在状态为 `ready` 后通过临时 loopback HTTP 服务执行单个 WPS Action。

当前里程碑开放清单中的只读和普通写入 WPS Action；破坏性操作需要明确确认并由 Runner 强制校验 `confirmed: true`。完整能力迁移由后续迁移 Issue 继续交付。下方 Node.js、MCP 和旧平台桥说明仍记录尚未删除的旧运行时，在替代路径通过完整验收前继续保留。

## 项目定位

本项目是MCP Server + Skills框架，让AI助手（Claude Code/Cursor/Augment等）能操控WPS Office。
- 235个MCP专业工具 + 12个内置工具 = 247个
- 支持Excel(82工具) / Word(32工具) / PPT(112工具) / 通用(9工具)
- 支持macOS、Windows、Linux

## 前提条件

- 已安装 WPS Office（https://www.wps.cn/ 或 https://www.wps.com/）
- Node.js >= 18.0.0
- Git

## 自动安装（AI执行）

以下步骤由AI助手自动执行。需要人类交互的步骤标注 ⚠️。

### 步骤1: 克隆项目

```bash
git clone https://github.com/lc2panda/wps-skills.git
cd wps-skills
```

如果项目已存在，跳过此步。

### 步骤2: 安装依赖并编译

```bash
cd wps-office-mcp
npm install
npm run build
cd ..
```

### 步骤3: 配置MCP Server

根据使用的AI工具，将以下配置写入对应文件。注意将路径替换为实际的项目绝对路径。

**Claude Code** — 写入 `~/.claude/settings.json`：
```json
{
  "mcpServers": {
    "wps-office": {
      "command": "node",
      "args": ["/你的路径/wps-skills/wps-office-mcp/dist/index.js"]
    }
  }
}
```

**Cursor** — 写入项目根目录 `.cursor/mcp.json`：
```json
{
  "mcpServers": {
    "wps-office": {
      "command": "node",
      "args": ["/你的路径/wps-skills/wps-office-mcp/dist/index.js"]
    }
  }
}
```

**OpenAI Codex CLI** — 写入 `~/.codex/config.toml`：
```toml
[mcp_servers.wps-office]
command = "node"
args = ["/你的路径/wps-skills/wps-office-mcp/dist/index.js"]
```
或命令行注册：`codex mcp add wps-office -- node /你的路径/wps-skills/wps-office-mcp/dist/index.js`

**Augment / 其他MCP兼容IDE** — 参考各IDE的MCP Server配置文档，使用相同的command和args。本项目 MCP Server 为标准 stdio 实现（spec 2025-11-25），与所有 MCP 一等客户端（Claude Code / Cursor / Codex CLI / GitHub Copilot CLI / Windsurf 等）兼容。

### 步骤4: 安装WPS加载项

⚠️ 需要人工操作（AI无法直接操作WPS应用）：

```bash
# macOS
bash scripts/auto-install-mac.sh

# Windows (PowerShell)
powershell scripts/install.ps1

# Linux
bash scripts/install.sh
```

⚠️ 安装后必须重启WPS Office才能生效。

### 步骤5: 安装Skills（仅Claude Code需要）

```bash
# 创建skills目录（如不存在）
mkdir -p ~/.claude/skills

# 创建符号链接
ln -sf "$(pwd)/skills/wps-excel" ~/.claude/skills/wps-excel
ln -sf "$(pwd)/skills/wps-word" ~/.claude/skills/wps-word
ln -sf "$(pwd)/skills/wps-ppt" ~/.claude/skills/wps-ppt
ln -sf "$(pwd)/skills/wps-office" ~/.claude/skills/wps-office
```

### 步骤6: 验证安装

```bash
# 验证MCP Server可启动
node wps-office-mcp/dist/index.js &
# 应看到 "MCP Server started successfully" 日志
kill %1 2>/dev/null
```

## 架构

```
Skills层(SKILL.md自然语言指导)
  ↓ Claude Code调用
MCP Server层(247个工具)
  ↓ wpsClient.executeMethod()
执行层
  ├── macOS: wps-claude-assistant (231 action, HTTP轮询)
  └── Windows: wps-com.ps1 (231 action, COM接口)
```

## 工具清单

| 应用 | 工具数 | 主要能力 |
|------|--------|---------|
| Excel | 82 | 公式/数据/图表/透视表/工作表/格式/工作簿/行列/批注保护/图片导出 |
| Word | 32 | 格式/内容/文档管理/页眉页脚/批注/模板填写/段落结构/校对修订 |
| PPT | 112 | 幻灯片/形状/图片/表格/美化/动画/图表/3D/数据可视化/图片导出 |
| 通用 | 9 | 保存/连接检测/文本选取/格式转换 |
| 内置 | 12 | 连接检查/万能方法调用/数据缓存 |

## 故障排除

| 问题 | 解决方案 |
|------|---------|
| MCP连接失败 | 确认 `npm install && npm run build` 已执行，检查dist/index.js存在 |
| WPS未响应 | 重启WPS Office，确认加载项已安装 |
| "arguments error" | 重新运行安装脚本，重启WPS |
| Linux找不到插件 | 查看INSTALL.md中的Linux专用指南 |
| 工具调用返回null | 确认WPS中已打开对应类型的文档 |

## 许可证

MIT
