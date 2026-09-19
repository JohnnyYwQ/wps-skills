# 安装 WPS Skills 插件

一个插件包含 `wps-word`、`wps-excel`、`wps-ppt` 三个 Skill，支持 Codex 和 Claude Code。安装包自带完整 Python 运行时源码、操作定义和 PowerShell 脚本，不需要项目仓库或第三方 Python 包。

## 环境要求

在实际操作 WPS 的 Windows 机器上安装和使用：Python 3.10+、Windows PowerShell 5.1、已安装相应 WPS 应用，并处于已登录的交互桌面。`python` 应指向 Windows Python；WSL/Linux Python 无法直接使用这套 Windows COM 执行端。

宿主需要支持插件安装、文件写入和本机脚本执行。Claude Code 在 Windows 上还需按其安装说明配置终端环境。插件不会代为安装 WPS、Python，也不会修改宿主的执行权限。

## 解压安装包

将 ZIP 解压到一个固定目录，例如 `C:/Tools/wps-skills`。这里应直接包含 `plugins/`、`.agents/`、`.claude-plugin/` 和本说明。复制时保留点号开头的目录。

以下命令均在 Windows 终端运行，路径替换为实际解压目录。仅需安装到使用的宿主；安装后保留解压目录，供插件更新和重新安装使用。

## Codex

```text
codex plugin marketplace add "C:/Tools/wps-skills"
codex plugin add wps-skills@wps-skills-local
```

这会登记安装包内的本地插件目录，并安装 `wps-skills`。新建对话后选择 `wps-word`、`wps-excel` 或 `wps-ppt`，描述需求即可。

## Claude Code

```text
claude plugin marketplace add "C:/Tools/wps-skills"
claude plugin install wps-skills@wps-skills-local --scope user
```

重新启动 Claude Code 后，可以直接描述办公需求，或使用 `/wps-skills:wps-word`、`/wps-skills:wps-excel`、`/wps-skills:wps-ppt` 选择相应能力。

临时加载同一插件目录也可以使用：

```text
claude --plugin-dir "C:/Tools/wps-skills/plugins/wps-skills"
```

插件目录和安装目录不同：`--plugin-dir` 指向内层 `plugins/wps-skills`，marketplace 命令指向外层解压目录。

## 使用与更新

先确认安装包中的查询脚本可启动，这一步不操作 WPS：

```text
python "C:/Tools/wps-skills/plugins/wps-skills/skills/wps-word/scripts/schema.py" createDocument writeContent saveAs
```

例如：“新建一份项目启动通知，保存为 C:/Documents/项目通知.docx。”输出路径应为实际绝对路径，父目录需存在。未要求保存时，文档留在 WPS 中，不自动保存。

如果此前已独立安装同名 Skill，迁移后只保留一种加载来源，避免宿主同时发现不同版本。升级时使用完整的新安装包，不将旧文件和新文件混合。固定安装目录更新后，Codex 重新运行 `codex plugin add wps-skills@wps-skills-local`；Claude Code 运行 `claude plugin update wps-skills@wps-skills-local`，随后新建会话。版本号由发布包统一管理。

更多操作说明见 [项目介绍](README.md)。安装方式参考 [Codex 官方文档](https://developers.openai.com/plugins/build/plugins) 和 [Claude Code 官方文档](https://code.claude.com/docs/en/plugin-marketplaces)。

## 从源码构建

在仓库根目录运行：

```sh
python scripts/build/plugin.py
```

输出目录为 `build/plugins/wps-skills/`，同级生成 `wps-skills.zip` 和 `wps-skills.zip.sha256`。安装包包含双宿主目录清单、插件清单、三个完整 Skill、中文说明及逐文件 SHA-256 清单。

已有输出不会被覆盖；再次构建时使用新的目录，例如：

```sh
python scripts/build/plugin.py --output build/plugins/wps-skills-next
```

外层输出目录名可以改变，内层插件名称始终为 `wps-skills`。源码中的 `src/main/resources/plugins/wps-skills/` 是构建模板，不能直接安装。

## 仅安装独立 Skill

不使用插件系统时，也可以按应用分别构建：

| Skill | WPS COM 注册 | 构建命令 |
| --- | --- | --- |
| wps-word | KWPS.Application | `python scripts/build/word.py --output build/skills/wps-word` |
| wps-excel | KET.Application | `python scripts/build/excel.py --output build/skills/wps-excel` |
| wps-ppt | KWPP.Application | `python scripts/build/ppt.py --output build/skills/wps-ppt` |

将整个 Skill 文件夹复制到宿主支持的 Skill 目录，保留 `scripts/`、`references/` 和 `runtime/`。每份独立包不依赖另两个 Skill。
