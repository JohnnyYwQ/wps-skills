<div align="center">

# WPS Skills

**简体中文** · [English](README.en.md)

**让 Agent 在本机真实 WPS 中完成办公任务**

从自然语言需求，到可编辑的 Word 文档、Excel 表格和 PPT 演示文稿。

<p>
  <code>Codex</code> &nbsp; <code>Claude Code</code> &nbsp; <code>Windows WPS</code> &nbsp; <code>Python 3.10+</code>
</p>

[功能概览](#功能概览) · [快速开始](#快速开始) · [使用示例](#使用示例) · [完整安装指南](INSTALL.md)

</div>

## 功能概览

一个插件，三个可独立使用的 Skill，共 **85 个操作（Action）**。支持创建、读取、编辑、排版和文件交付，也支持按应用单独安装。

| 应用 | 你可以做什么 | 保存与导出 |
| --- | --- | --- |
| 📄 **[Word](src/main/resources/skills/wps-word/SKILL.md)** · 14 个操作 | 编写正文、查找替换、调整排版，插入表格与图片，设置页眉页脚和页面布局 | DOCX · PDF |
| 📊 **[Excel](src/main/resources/skills/wps-excel/SKILL.md)** · 34 个操作 | 管理工作表、读写数据、计算公式，设置区域格式，调整行列、排序与筛选 | XLSX · PDF |
| 📽️ **[PPT](src/main/resources/skills/wps-ppt/SKILL.md)** · 37 个操作 | 编排幻灯片、编辑文字与形状，插入图片与表格，调整布局和讲者备注 | PPTX · PDF · PNG |

Agent 负责理解需求和安排操作，执行端通过 Windows WPS COM 操作真实文档，并返回执行结果。运行时使用 Python 标准库和 Windows PowerShell，**无第三方 Python 包依赖**。

## 快速开始

**准备环境：** Windows、Python 3.10+、Windows PowerShell 5.1，以及已安装并可正常打开的相应 WPS 应用。实际操作需要在已登录的交互桌面中执行，宿主需支持插件、文件写入和本机脚本调用。

### 1. 构建与解压

在仓库根目录运行：

```sh
python scripts/build/plugin.py
```

将生成的 `build/plugins/wps-skills.zip` 解压到固定目录，例如 `C:/Tools/wps-skills`。解压后的结构如下，点号开头的目录也需保留：

```text
C:/Tools/wps-skills/
├── .agents/           Codex 安装目录清单
├── .claude-plugin/    Claude Code 安装目录清单
├── plugins/
│   └── wps-skills/    插件与三个完整 Skill
└── README.md          安装说明
```

### 2. 安装到你的宿主

在 Windows 终端中运行对应命令，路径替换为实际解压目录。

<details open>
<summary><strong>Codex</strong></summary>

```text
codex plugin marketplace add "C:/Tools/wps-skills"
codex plugin add wps-skills@wps-skills-local
```

</details>

<details open>
<summary><strong>Claude Code</strong></summary>

```text
claude plugin marketplace add "C:/Tools/wps-skills"
claude plugin install wps-skills@wps-skills-local --scope user
```

</details>

### 3. 开始使用

新建会话，选择 `wps-word`、`wps-excel` 或 `wps-ppt`，直接描述目标、内容和输出位置。Claude Code 也可通过以下命令选择能力：

```text
/wps-skills:wps-word
/wps-skills:wps-excel
/wps-skills:wps-ppt
```

升级、临时加载和独立 Skill 安装方式见 **[完整安装指南 →](INSTALL.md)**。

## 使用示例

**📄 Word · 项目启动通知**

> 新建一份“星河知识库”项目启动通知，包含项目目标、参与部门和四周实施安排，排版整齐，保存为 `C:/Documents/项目通知.docx`。

**📊 Excel · 项目预算表**

> 新建项目预算表，列出人员、设备和服务三类费用，用公式计算合计，设置金额格式，保存为 `C:/Documents/项目预算.xlsx`。

**📽️ PPT · 项目汇报**

> 制作一份四页的项目汇报，包含项目背景、实施计划、当前进展和下一步安排，保存为 `C:/Documents/项目汇报.pptx`。

输出路径使用本机实际绝对路径，父目录需已存在。**仅在明确请求时保存**；导出 PDF 或 PNG 不代表已保存源文档。

## 如何完成任务

**理解需求 → 按需查询操作定义 → 提交多步计划 → 操作 WPS 并验证 → 报告结果**

| 设计 | 作用 |
| --- | --- |
| **按需提供定义** | 先通过 Skill 了解操作用途，再查询本次需要的参数和结果结构，减少无关内容进入上下文。 |
| **多步任务编排** | 用结果引用衔接操作依赖，在一次提交内传递中间结果，复用文档绑定与通信进程。 |
| **精确文档绑定** | 操作指向明确文档，并校验相关内容的版本或 token，降低误改窗口和使用过期定位的风险。 |
| **回读验证与回执** | 按操作契约验证效果，记录每步结果；响应丢失后可查询，Task 结束后释放所属执行资源。 |

一个 Task 处理一个应用中的一份文档；跨应用协作或需要先读取再决定内容时，由 Agent 分别提交。具体参数和能力边界见各 Skill 的说明及操作定义。

## 进阶使用

<details>
<summary><strong>手动查询、提交 Task 与读取回执</strong></summary>

以下以解压后的 Word Skill 为例，在 Windows 终端执行。先进入脚本目录，后续命令使用同一工作目录：

```text
cd "C:/Tools/wps-skills/plugins/wps-skills/skills/wps-word/scripts"
```

**查询定义**——获取所需参数和结果结构，不启动 WPS：

```text
python schema.py createDocument writeContent saveAs
```

**提交任务**——按照 Skill 说明准备 UTF-8 JSON 文件，每份新请求使用新路径：

```text
python word.py --app word --task-file "C:/Tasks/request-001.json"
```

**查询回执**——命令输出丢失或超时时，使用原请求路径：

```text
python word.py --app word --task-status-file "C:/Tasks/request-001.json"
```

即使输入文件已被消费，查询仍使用原路径。重复提交已受理路径不会重放操作。执行失败或结果不确定时，后续步骤停止，已发生的效果保留，不自动回滚或重试。

Excel、PPT 使用各自 Skill 目录中的 `schema.py` 和 `excel.py`、`ppt.py`，并指定对应的 `--app`。请求格式与完整示例见各包的 `SKILL.md`。

</details>

<details>
<summary><strong>换机前检查 Windows 环境</strong></summary>

在 Windows 上，从仓库根目录运行 doctor：

```sh
python scripts/doctor.py --app word
python scripts/doctor.py --app excel
python scripts/doctor.py --app ppt
```

doctor 默认检查桌面会话、COM 注册、PowerShell 和中文通信等条件，不启动 WPS 或修改文档。环境检查通过不等于所有文档操作都可用，更多选项运行 `python scripts/doctor.py --help` 查看。

</details>

<p align="center">
  <a href="INSTALL.md">安装指南</a> ·
  <a href="LICENSE">MIT License</a>
</p>
