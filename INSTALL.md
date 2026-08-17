# WPS Skills 安装指南

> **本文件面向 AI 编程助手**（Claude Code、Cursor 等）。读取本文件后，你应能完全理解安装流程并自动执行所有步骤。按顺序执行，不要跳步。

---

## 安装前状态检查

在执行安装前，先判断是否已经安装过。依次执行以下检查：

```bash
# 1. 检查 MCP Server 是否已注册
claude mcp list 2>/dev/null | grep -i wps

# 2. 检查 Skills 是否已注册
ls ~/.claude/skills/wps-excel/SKILL.md 2>/dev/null

# 3. 检查编译产物是否存在
ls wps-office-mcp/dist/index.js 2>/dev/null
```

**判断逻辑**：
- 三项全部存在 → 已安装，无需重复执行。告知用户"WPS Skills 已安装，重启 Claude Code 和 WPS Office 即可使用"
- 部分存在 → 从缺失的步骤开始补装
- 全部不存在 → 执行完整安装流程

---

## macOS 安装

### Step 1: 检测前置条件

```bash
# 检测操作系统（必须是 Darwin）
uname -s

# 检测 WPS Office 是否已安装
ls /Applications/wpsoffice.app

# 检测 Node.js 版本（必须 >= 18.0.0）
node -v
```

**如果前置条件不满足**：
- WPS Office 不存在 → 告知用户安装：https://www.wps.cn/product/wpsmac
- Node.js 不存在或版本 < 18 → 告知用户安装：https://nodejs.org/
- 前置条件不满足时，停止安装流程，不要继续

### Step 2: 运行一键安装脚本

```bash
# 进入项目根目录（即包含本文件的目录）
cd <项目根目录>

# 赋予执行权限并运行
chmod +x scripts/auto-install-mac.sh
./scripts/auto-install-mac.sh
```

脚本自动完成以下全部操作：
- 二次检测前置条件（macOS、WPS Office、Node.js 18+）
- 创建 WPS 加载项目录（如不存在）
- 安装加载项文件到 `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/claude-assistant_/`（目录名尾部 `_` 是必须的）
- 更新 publish.xml 注册加载项
- 执行 `npm install` 安装依赖
- 执行 `npm run build` 编译 TypeScript
- 执行 `claude mcp add wps-office` 注册 MCP Server
- 创建软链接注册 Skills 到 `~/.claude/skills/`

### Step 3: 安装后验证

逐项执行以下检查，全部通过才算安装成功：

```bash
# 1. MCP Server 已注册（输出应包含 wps-office）
claude mcp list

# 2. Skills 已注册（应列出 wps-excel, wps-word, wps-ppt, wps-office 四个目录/链接）
ls ~/.claude/skills/

# 3. 编译产物存在
ls wps-office-mcp/dist/index.js

# 4. WPS 加载项已安装（目录名以 _ 结尾）
ls ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/claude-assistant_/

# 5. publish.xml 已注册加载项（输出应包含 claude-assistant）
grep "claude-assistant" ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml
```

### Step 4: 告知用户

安装完成后，必须告知用户执行以下操作：
1. **重启 Claude Code**（必须！Skills 只在启动时加载）
2. 重启 WPS Office
3. 打开任意 Excel/Word/PPT 文档
4. 在 WPS 顶部查看 "Claude助手" 选项卡，确认状态显示 "轮询中"

### ⚠️ 已知问题

**macOS 沙盒目录访问权限弹窗**
- 现象：首次执行验证命令访问 `~/Library/Containers/com.kingsoft.wpsoffice.mac/...` 时，macOS 会弹出文件访问权限请求
- 处理：点击「允许」即可，之后不再弹出
- 影响：不影响安装结果，属正常 macOS 沙盒机制

---

## Linux 安装

### Step 1: 检测前置条件

```bash
# 检测操作系统（必须是 Linux）
uname -s

# 检测 WPS Office 是否已安装
which wps || ls /opt/kingsoft/wps-office

# 检测 Node.js 版本（必须 >= 18.0.0）
node -v
```

**如果前置条件不满足**：
- WPS Office 不存在 -> 告知用户安装：https://linux.wps.cn
- Node.js 不存在或版本 < 18 -> 告知用户安装：https://nodejs.org/

### Step 2: 手动安装

```bash
# 进入项目根目录
cd <项目根目录>

# 安装依赖并编译
cd wps-office-mcp
npm install
rm -rf dist
npm run build
cd ..

# 复制加载项到 WPS 目录（目录名必须以 _ 结尾）
mkdir -p ~/.local/share/Kingsoft/wps/jsaddons
cp -R wps-claude-assistant ~/.local/share/Kingsoft/wps/jsaddons/claude-assistant_

# 创建 publish.xml
# 注：enable="enable_dev" 为开发模式（默认）；若加载失败可改为 enable="true"（发布模式）
cat > ~/.local/share/Kingsoft/wps/jsaddons/publish.xml << 'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<jsplugins>
  <jsplugin name="claude-assistant" type="wps,et,wpp" url="claude-assistant_/" enable="enable_dev"/>
</jsplugins>
EOF

# 注册 MCP Server
claude mcp add wps-office node $(pwd)/wps-office-mcp/dist/index.js

# 注册 Skills
mkdir -p ~/.claude/skills
ln -sf $(pwd)/skills/wps-excel ~/.claude/skills/wps-excel
ln -sf $(pwd)/skills/wps-word ~/.claude/skills/wps-word
ln -sf $(pwd)/skills/wps-ppt ~/.claude/skills/wps-ppt
ln -sf $(pwd)/skills/wps-office ~/.claude/skills/wps-office
```

### Step 3: 安装后验证

```bash
# 1. MCP Server 已注册
claude mcp list

# 2. Skills 已注册
ls ~/.claude/skills/

# 3. 编译产物存在
ls wps-office-mcp/dist/index.js

# 4. WPS 加载项已安装（目录名以 _ 结尾）
ls ~/.local/share/Kingsoft/wps/jsaddons/claude-assistant_/

# 5. publish.xml 已注册
grep "claude-assistant" ~/.local/share/Kingsoft/wps/jsaddons/publish.xml
```

### Step 4: 告知用户

1. **重启 Claude Code**（必须！）
2. 重启 WPS Office
3. 打开任意文档，查看 "Claude助手" 选项卡

### ⚠️ Linux 启动顺序（重要）

Issue #17 反馈：在 Kylin Linux 等发行版上，若 WPS 已运行后再启动 Claude Code（MCP Server），WPS 进程可能异常退出。

**正确启动顺序**：
1. 先启动 **Claude Code**（含 MCP Server 初始化）
2. 再启动 **WPS Office**（加载项 HTTP 轮询会找到已就绪的 58891 端口）
3. 使用 WPS MCP 工具

若 WPS 已在运行，建议先 `pkill -9 wps && pkill -9 wpp && pkill -9 et`，再按上述顺序重启。

### Linux 关键路径参考

| 项目 | 路径 |
|------|------|
| WPS 加载项基础目录 | `~/.local/share/Kingsoft/wps/jsaddons/` |
| 加载项安装目录 | `<基础目录>/claude-assistant_/`（尾部 `_` 必须） |
| publish.xml | `<基础目录>/publish.xml` |

---

## Windows 安装

### Step 1: 检测前置条件

```powershell
# 检测 WPS Office 加载项目录是否存在
Test-Path "$env:APPDATA\kingsoft\wps\jsaddons"

# 检测 Node.js 版本（必须 >= 18.0.0）
node -v
```

**如果前置条件不满足**：
- WPS 加载项目录不存在 → 告知用户安装 WPS Office：https://www.wps.cn/
- Node.js 不存在或版本 < 18 → 告知用户安装：https://nodejs.org/
- 前置条件不满足时，停止安装流程，不要继续

### Step 2: 运行一键安装脚本

```powershell
# 进入项目根目录
cd <项目根目录>

# 执行安装脚本
powershell -ExecutionPolicy Bypass -File scripts/auto-install.ps1
```

脚本自动完成以下全部操作：
- 检测 Node.js 18+ 版本
- 执行 `npm install` 安装依赖
- 执行 `npm run build` 编译 TypeScript
- 配置 Claude Code MCP（写入 `%USERPROFILE%\.claude\settings.json`）
- 复制 Skills 到 `%USERPROFILE%\.claude\skills\`
- 安装 WPS 加载项到 `%APPDATA%\kingsoft\wps\jsaddons\wps-claude-addon_\`（目录名尾部 `_` 是必须的）
- 更新 publish.xml 注册加载项

### Step 3: 安装后验证

```powershell
# 1. MCP Server 已注册
claude mcp list

# 2. Skills 已注册（应列出 wps-excel, wps-word, wps-ppt, wps-office）
Get-ChildItem "$env:USERPROFILE\.claude\skills"

# 3. 编译产物存在
Test-Path "wps-office-mcp\dist\index.js"

# 4. WPS 加载项已安装
Test-Path "$env:APPDATA\kingsoft\wps\jsaddons\wps-claude-addon_"

# 5. publish.xml 已注册加载项
Select-String -Path "$env:APPDATA\kingsoft\wps\jsaddons\publish.xml" -Pattern "wps-claude-addon"
```

### Step 4: 告知用户

安装完成后，必须告知用户：
1. **重启 Claude Code**（必须！）
2. 重启 WPS Office
3. 查看 "Claude助手" 选项卡

---

## 关键路径参考

### macOS

| 项目 | 路径 |
|------|------|
| WPS 加载项基础目录 | `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/` |
| 加载项安装目录 | `<基础目录>/claude-assistant_/`（尾部 `_` 必须） |
| publish.xml | `<基础目录>/publish.xml` |
| Skills 注册目录 | `~/.claude/skills/`（4 个软链接） |
| MCP Server 入口 | `<项目根目录>/wps-office-mcp/dist/index.js` |
| HTTP 轮询端口 | `58891` |

### Windows

| 项目 | 路径 |
|------|------|
| WPS 加载项基础目录 | `%APPDATA%\kingsoft\wps\jsaddons\` |
| 加载项安装目录 | `<基础目录>\wps-claude-addon_\`（尾部 `_` 必须） |
| publish.xml | `<基础目录>\publish.xml` |
| Skills 注册目录 | `%USERPROFILE%\.claude\skills\`（复制方式，非软链接） |
| MCP Server 配置 | `%USERPROFILE%\.claude\settings.json` |

---

## 错误处理

遇到安装错误时，按以下对照表处理：

### npm install 失败

```bash
# 清除缓存后重试
cd wps-office-mcp
rm -rf node_modules package-lock.json
npm install
```

如果仍失败，检查 Node.js 版本：
```bash
node -v
# 必须 >= 18.0.0，否则升级 Node.js
```

### npm run build（TypeScript 编译）失败

```bash
cd wps-office-mcp
rm -rf dist node_modules
npm install
npm run build
```

如果报 `tsc: command not found`，说明 typescript 未安装为依赖，检查 package.json 的 devDependencies 中是否有 typescript。

### MCP Server 注册失败

手动注册：
```bash
claude mcp add wps-office node <项目根目录的绝对路径>/wps-office-mcp/dist/index.js
```

注意：`<项目根目录的绝对路径>` 必须替换为实际路径，不能使用相对路径或变量。

### Skills 软链接创建失败

```bash
PROJECT_DIR=<项目根目录的绝对路径>
mkdir -p ~/.claude/skills
ln -sf "$PROJECT_DIR/skills/wps-excel" ~/.claude/skills/wps-excel
ln -sf "$PROJECT_DIR/skills/wps-word" ~/.claude/skills/wps-word
ln -sf "$PROJECT_DIR/skills/wps-ppt" ~/.claude/skills/wps-ppt
ln -sf "$PROJECT_DIR/skills/wps-office" ~/.claude/skills/wps-office
```

验证：
```bash
ls -la ~/.claude/skills/
# 应看到 4 个软链接，指向项目中的 skills/ 子目录
```

### WPS 加载项未显示 "Claude助手" 选项卡

1. 确认加载项目录已正确复制且名称以 `_` 结尾：
```bash
# macOS
ls ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/claude-assistant_/
# 应包含 main.js, manifest.xml, ribbon.xml 等文件
```

2. 确认 publish.xml 包含注册条目：
```bash
# macOS
cat ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml
# 应包含 <jsplugin name="claude-assistant" .../> 条目
```

3. 清理 WPS 实际加载目录的 Gatekeeper 隔离属性：
```bash
xattr -dr com.apple.quarantine ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/claude-assistant_/
```

4. 强制退出并重启 WPS：
```bash
# macOS
pkill -f wpsoffice
sleep 2
open /Applications/wpsoffice.app
```

### HTTP 轮询端口 58891 被占用（macOS）

```bash
# 查看端口占用
lsof -i :58891

# 终止占用进程
kill <PID>
```

### macOS 加载项目录权限不足

```bash
# 手动创建目录
mkdir -p ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons

# 修复权限
chmod -R 755 ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons
```

---

## 已知问题与故障排除（GitHub Issues）

### Issue #6: MCP 连接失败（"Failed to connect"）

**现象**：`claude mcp list` 显示 `wps-office: Failed to connect`。

**根因**：dist 目录中的编译产物过期，存在工具名称重复注册导致 MCP Server 启动即崩溃。

**解决方案**：
```bash
cd wps-office-mcp
# 清除旧编译产物后重新编译（关键步骤！）
rm -rf dist
npm run build
# 验证启动是否正常（应看到 "Server started successfully"）
node dist/index.js 2>&1 | head -5
# Ctrl+C 退出
```

如果仍然失败，完整重建：
```bash
cd wps-office-mcp
rm -rf dist node_modules
npm install
npm run build
```

然后重新注册 MCP：
```bash
claude mcp remove wps-office
claude mcp add wps-office node <项目绝对路径>/wps-office-mcp/dist/index.js
```

### Issue #5: Linux 平台 WPS 找不到加载项

**现象**：Linux（如 ArchLinux）安装成功但 WPS 中看不到 Claude 助手选项卡。

**根因**：安装脚本中 Linux 加载项目录名缺少尾部 `_` 后缀，且未更新 publish.xml。WPS jsaddons 规范要求目录名以 `_` 结尾才能被识别。

**解决方案**：
```bash
# 1. 检查当前安装路径（错误的旧路径）
ls ~/.local/share/Kingsoft/wps/jsaddons/wps-claude-addon 2>/dev/null

# 2. 如果存在旧目录，删除并重新安装
rm -rf ~/.local/share/Kingsoft/wps/jsaddons/wps-claude-addon

# 3. 手动复制到正确路径（目录名必须以 _ 结尾）
cp -R <项目根目录>/wps-claude-assistant ~/.local/share/Kingsoft/wps/jsaddons/claude-assistant_

# 4. 创建 publish.xml（Linux 上也需要）
cat > ~/.local/share/Kingsoft/wps/jsaddons/publish.xml << 'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<jsplugins>
  <jsplugin name="claude-assistant" type="wps,et,wpp" url="claude-assistant_/" enable="enable_dev"/>
</jsplugins>
EOF

# 5. 重启 WPS
pkill -f wps
# 然后重新打开 WPS
```

**注意**：不同 Linux 发行版 WPS 加载项目录可能不同。已知路径：
- 通用：`~/.local/share/Kingsoft/wps/jsaddons/`
- 部分发行版：`~/.kingsoft/wps/jsaddons/`

如果以上路径都不生效，可通过以下方式查找：
```bash
find / -path "*/Kingsoft/wps/jsaddons" -type d 2>/dev/null
find / -path "*kingsoft/wps/jsaddons" -type d 2>/dev/null
```

### Issue #4: WPS 加载项启动报 "arguments error"

**现象**：WPS 弹出 `ERROR: arguments error at <anonymous>:1:89`。

**根因**：manifest.xml 中缺少 `<ribbon>` 和 `<scripts>` 标签声明，导致 WPS 无法正确解析加载项入口。

**解决方案**：

已修复 `wps-claude-assistant/manifest.xml`，确保包含完整的 ribbon 和 scripts 声明。如果仍遇到此错误：

```bash
# 1. 检查 manifest.xml 是否包含 ribbon 和 scripts 声明
grep -E "ribbon|scripts" <加载项安装目录>/manifest.xml
# 应看到类似：
#   <ribbon src="ribbon.xml"/>
#   <script src="main.js"/>

# 2. 确认 ribbon.xml 和 main.js 文件存在
ls <加载项安装目录>/ribbon.xml
ls <加载项安装目录>/main.js

# 3. 重新复制加载项（使用最新版本）
# macOS:
rm -rf ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/claude-assistant_
cp -R <项目根目录>/wps-claude-assistant ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/claude-assistant_

# 4. 重启 WPS
pkill -f wpsoffice
sleep 2
open /Applications/wpsoffice.app
```
