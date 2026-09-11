# WPS 自动化基础

[English](README.md) | [简体中文](README.zh-CN.md)

本项目为智能体提供 WPS 自动化能力，采用独立 Application Skill、明确的 Action Contracts 和单文档 Action Session 架构。目前已接入 Windows WPS Writer、表格和演示，分别提供可独立安装的 `wps-word`、`wps-excel` 和 `wps-ppt` Skill。

智能体可以通过 Skill 了解可用能力，再用 Python Session Client 创建或打开文档、执行编辑、验证结果并按需保存。Python Session Host 负责请求校验和会话管理，PowerShell bridge 通过 WPS COM 操作同一个实际文档。

## PPT

PPT 已提供 **37 个原生动作**，覆盖新建和已有 `.pptx` 的读取、编辑、排版与保存。
本次新增 18 项常用能力：

| 类别 | 新动作 |
| --- | --- |
| 形状样式与排版 | getShapeStyle、formatShape、renameShape、setShapeOrder、alignShapes、distributeShapes |
| 段落与文本框 | formatParagraph、setTextBoxLayout |
| 背景、隐藏和页名称 | getSlideSettings、setSlideSettings |
| 演讲备注 | getSlideNotes、setSlideNotes |
| 查找替换 | findText、replaceText |
| 图片与表格 | addImage、addTable、readTable、writeTable |

每次编辑使用对应读取结果的 token 并返回实际读回内容；一场会话始终绑定一个确切文稿。
需要 Windows 和注册为 `KWPP.Application` 的 WPS 演示。结构操作最多 200 页，
单页最多读取 100 个顶层形状，单形状文字最多 10000 UTF-16 单元。
表格最多 100 格，图片只嵌入 PNG/JPEG；具体上限和 token 来源见 Skill 参考。
已支持新建文稿、首次保存、另存为、PDF 导出和单页 PNG 导出；图表和动画尚未提供。
原生形状复制尚未通过验证，不在正式能力集内。

```powershell
python scripts/build/ppt.py --output build/skills/wps-ppt
python build/skills/wps-ppt/scripts/ppt.py --app ppt --index
python scripts/call.py --app ppt --resolve openPresentation listSlides getSlideInfo setShapeText save
```

在已同步的 Windows 测试机上，从**普通桌面 PowerShell** 粘贴下面完整命令。
它复制空白模板，所有可见内容编辑都通过正式 PPT 会话完成，最后保存并保留 WPS 窗口；
不会额外弹出 cmd/PowerShell 黑窗口：

```powershell
$repo = Join-Path $env:LOCALAPPDATA 'Temp\wps-ppt-20260905'
$python = Join-Path $env:USERPROFILE '.conda\envs\yolov8app\python.exe'
$launcher = Join-Path $repo 'src\main\resources\wps_skills\ppt\windows\demo_launcher.ps1'
& ([scriptblock]::Create([IO.File]::ReadAllText($launcher))) -RepositoryRoot $repo -PythonPath $python -Delay 1.5
```

其他机器将 `$repo`、`$python` 改为实际仓库和 Python 路径。
允许执行 `.ps1` 时也可在仓库根目录运行 `./scripts/demo/ppt.ps1 -PythonPath <python.exe绝对路径>`。
演示模板复制不是正式的新建文稿动作。输出目录和文件名每次唯一，报告随文件保存。

真机验收会生成自己的测试文稿：

```powershell
python scripts/validate/ppt.py --output-dir build/ppt-acceptance
python scripts/validate/ppt_common.py --output-dir build/ppt-common-acceptance
```

输出目录必须不存在。参见 [PPT Skill](src/main/resources/skills/wps-ppt/SKILL.md)
及[原生能力证据](src/test/resources/wps_skills/ppt/type_library/EVIDENCE.md)。

## Excel

支持新建与已有 `.xlsx` 工作簿的 34 个 Actions：

| 类别 | Actions |
| --- | --- |
| 工作簿 | `openWorkbook`、`getWorkbookInfo`、`save` |
| 工作表 | `listWorksheets`、`getWorksheetInfo`、`addWorksheet`、`renameWorksheet`、`copyWorksheet`、`moveWorksheet`、`deleteWorksheet` |
| 数据与公式 | `readRange`、`writeRange`、`setFormulas`、`calculateRange`、`clearRange`、`copyRange`、`findInRange`、`replaceInRange` |
| 排序筛选 | `sortRange`、`filterRange`、`clearFilter` |
| 格式与尺寸 | `formatRange`、`mergeRange`、`unmergeRange`、`setRowHeight`、`setColumnWidth`、`autoFitColumns` |
| 行列结构 | `insertRows`、`deleteRows`、`insertColumns`、`deleteColumns` |

每次指定准确工作表名称和最多 1000 个单元格的 A1 矩形。区域修改先 `readRange`，
工作表和行列结构修改先 `getWorksheetInfo`，使用对应 token 并回读验证。格式支持
字号、粗斜体、颜色、对齐、换行和数字格式，公式支持常用统计、查找和文本函数。
原位保存期间持续保护文档占用。已开放新建工作簿、首次保存、另存为和 PDF 导出；图表与透视表尚未开放。
结构操作也要求整表已用区域不超过 1000 个单元格。细节见 Excel Skill 的参考文档。

```bash
python scripts/build/excel.py --output build/skills/wps-excel
python build/skills/wps-excel/scripts/excel.py --app excel --index
python scripts/call.py --app excel --resolve openWorkbook readRange writeRange save
```

运行要求为 Windows、Python 3.8+、原生 Windows PowerShell 和注册为 `KET.Application`
的 WPS 表格。已在 WPS 12.0.0.28505 上真机验收。参见 [Excel Skill](src/main/resources/skills/wps-excel/SKILL.md)
和[能力验证记录](src/test/resources/wps_skills/excel/type_library/EVIDENCE.md)。

在 Windows 桌面的**普通 PowerShell（非管理员）**中运行演示，会打开 WPS 窗口，展示填写、公式和格式，保存后保留窗口。无控制台启动器在创建 Python 进程时设置 `CreateNoWindow`，将进度和错误转发到现有终端；不经额外 `cmd` 窗口。下面是本次 Windows 测试目录和 Python 环境对应的完整命令，其他安装位置需替换前面两个路径：

```powershell
$ErrorActionPreference = 'Stop'
$repo = Join-Path $env:LOCALAPPDATA 'Temp\wps-excel-20260905'
$python = Join-Path $env:USERPROFILE '.conda\envs\yolov8app\python.exe'
$launcher = Join-Path $repo 'src\main\resources\wps_skills\excel\windows\demo_launcher.ps1'

& ([scriptblock]::Create([IO.File]::ReadAllText($launcher))) `
    -RepositoryRoot $repo -PythonPath $python -Delay 2
```

如果运行环境允许执行 `.ps1` 文件，也可在仓库根目录运行 `./scripts/demo/excel.ps1 -PythonPath <python.exe绝对路径> -Delay 2`。Python 入口 `scripts/demo/excel.py` 仍可直接运行。每次演示使用唯一工作簿名，可以连续运行并保留之前的窗口。SSH 的非交互会话不能直接用于观察桌面演示。

Windows 真机验收使用一个尚不存在的输出目录，仅修改脚本生成的测试工作簿，结束后保留打开：

```powershell
python scripts/validate/excel.py --output-dir build/excel-acceptance --wps-version 12.0.0.28505
python scripts/validate/excel_common.py --output-dir build/excel-common-acceptance
```

## 当前能力

Word 正式 Application Contract Set 包含 14 个可执行 Action：

| Action | 作用 |
| --- | --- |
| `createDocument` | 创建一个新的未保存文档 |
| `openDocument` | 打开或复用指定路径的已有 `.docx` 文档 |
| `writeContent` | 写入结构化正文、标题及文字和段落格式 |
| `inspectDocument` | 读取内容、格式、结构、页面设置和文档状态 |
| `findContent` | 在指定范围内查找文字 |
| `replaceContent` | 按范围或有匹配数量约束的查询替换内容 |
| `insertTable` | 插入表格 |
| `insertImage` | 插入图片 |
| `setHeaderFooter` | 设置页眉页脚 |
| `setPageLayout` | 设置页面布局 |
| `insertBreak` | 插入分页符或分节符 |
| `save` | 将绑定文档保存到当前路径 |
| `saveAs` | 首次保存或另存为尚不存在的 DOCX |
| `exportPdf` | 导出 PDF |

支持分别设置西文字体与东亚字体，并在操作后读回验证。内容范围带有 Content Revision，避免后续操作误用文档修改前的旧位置。

`saveAs` 已进入三端正式能力集：首次保存和另存为都保留同一个实际文档，后续 `save` 使用新路径。只支持 `overwritePolicy: "failIfExists"`，不会覆盖已有目标。Excel 和 PPT 分别通过 `--app excel`、`--app ppt` 接入正式 CLI。

三端共 85 项 Action（Word 14、Excel 34、PPT 37）。保存／导出的观察验证有界：Excel 最多 20 张表、每张 UsedRange 最多 1000 格；PPT 最多 200 页、每页 100 个顶层形状。不宣称完整验证未支持的文档特性。使用 `python scripts/validate/persistence.py --output-dir build/persistence-acceptance` 运行三端持久化原生验收。

## 环境要求

- Python 3.8 或更高版本。
- 执行文档操作需要 Windows，以及已注册 `KWPS.Application` 的 WPS Writer。
- 使用系统原生 Windows PowerShell：`%WINDIR%\System32\WindowsPowerShell\v1.0\powershell.exe`。
- 无需第三方 Python 包，也无需安装常驻 Host 服务。

能力查询、Skill 构建和本地测试也可在 macOS/Linux 上运行。实际文档操作必须在 WPS 所在的 Windows 主机执行。

## 构建与使用 Word Skill

在仓库根目录运行，将 `<输出目录>` 替换为实际目标目录：

```bash
python scripts/build/word.py --output "<输出目录>/wps-word"
python "<输出目录>/wps-word/scripts/word.py" --app word --index
python "<输出目录>/wps-word/scripts/word.py" --app word --resolve createDocument writeContent inspectDocument
```

构建会在指定位置创建 `wps-word/` 目录，包含：

```text
wps-word/
  SKILL.md
  agents/openai.yaml
  references/
  scripts/word.py
  runtime/
    files.sha256.json
    src/main/python/wps_skills/
    src/main/resources/wps_skills/word/windows/
```

将完整的 `wps-word` 目录复制到目标智能体支持的技能目录。部署后的 Skill 使用随包 Runtime，无需仓库源码；只复制 `SKILL.md` 无法执行文档操作。

构建过程生成文件 SHA-256 清单，并拒绝覆盖已有目标目录。再次构建时，可用 `--output <新目录>/wps-word` 指定新的输出位置。源码是唯一维护位置，构建产物无需单独修改。

`--index` 返回由正式契约生成的紧凑 Action Index；`--resolve` 一次返回所需 Actions 的完整契约。这两个命令不会启动 Session、PowerShell 或 WPS。批量解析结果为 `partial` 或 `failed` 时，先调整操作计划，再执行文档修改。

阅读 [Word Skill](src/main/resources/skills/wps-word/SKILL.md) 了解完整工作流程。[会话使用说明](src/main/resources/skills/wps-word/references/session.md) 提供包内 `--start`、`--call`、`--status`、`--close` 命令和 JSON 参数示例；Agent 无需编写任务脚本，多次独立命令复用同一个 Session，并在每次响应后决定下一步。

源码中的 `src/main/resources/skills/wps-word/scripts/word.py` 也可以直接使用。更多安装说明见 [INSTALL.md](INSTALL.md)。

## 直接启动 Word Session

在安装了 WPS 的 Windows 主机上，从仓库根目录运行：

```powershell
python scripts/call.py --session --app word
```

Host 输出 `session.ready` 后等待单行 JSON 请求。例如：

```json
{"address":{"app":"word","action":"createDocument"},"params":{}}
```

每次收到完整 Action Response 后才能发送下一个请求。一个 Session 始终绑定同一个实际文档并复用一条 bridge，不通过当前活动窗口或 UI 焦点选择后续操作目标。结束时发送：

```json
{"control":"close"}
```

日常任务优先使用 Skill 提供的 Python Session Client。它处理请求顺序、超时、终止通知和真实错误，避免调用方重复实现通信逻辑。收到 `session.closing` 后应停止提交新请求，但继续读取后续 Action Response 和最终关闭记录；操作结果不确定时不得自动重放。

需要在远程桌面观察 WPS 时，应在用户已登录的桌面会话中启动。仅通过 SSH 执行命令不能保证窗口出现在远程桌面；远程启动方式取决于执行环境，Skill 不内置主机地址、账号或计划任务。

## 文档保存与清理

普通 Session 结束时释放自动化资源，保留文档打开，不会隐式保存。修改已有文件通常需要显式执行 `save`；用户要求保持未保存状态时除外。新建且没有输出路径的文档可以保留未保存状态，但必须在结果中明确说明。

仅对本次创建的可丢弃测试文档，可显式启用测试清理：

```powershell
python scripts/call.py --session --app word --debug-close-created-document
```

该选项只丢弃并关闭当前 Session 创建的文档，不关闭通过 `openDocument` 获取的文档，也不是 Word Action。需要保留内容的用户任务不应使用它。

Session 使用 Windows Job Object 管理自有桥接进程，使用 Guard、Document Lease 和 Document Quarantine 协调跨进程文档访问。正常启动会显示 WPS 文档窗口，隐藏桥接控制台；新建的 WPS 应用窗口使用普通窗口状态。对于异常小的窗口，会调整到合理的可见范围。

## 日志与耗时

Action Response 的 `traceLog` 指向该次操作的日志，其中记录 `elapsedMs`。Session 日志末尾记录：

- `sessionElapsedMs`：会话总耗时。
- `actionExecutionElapsedMs`：各 Action 执行耗时之和。
- `cleanupElapsedMs`：清理耗时。
- `actionCount`：执行的 Action 数量。

这些是诊断字段，不会扩展 Protocol v1 的 Action Response。进程墙钟时间、调用往返时间与 Action 执行时间应分别理解。Session 清理成功也不代表用户的文档任务已经完成，仍需检查编辑、验证和保存结果。

## 运行测试

在 macOS/Linux 的仓库根目录运行：

```bash
PYTHONPATH=src/main/python python -m unittest discover -s src/test/python -p 'test_*.py'
```

Windows PowerShell：

```powershell
$env:PYTHONPATH = "src/main/python"
python -m unittest discover -s src/test/python -p "test_*.py"
```

本地测试覆盖 fake 实现、真实子进程通信以及迁移到独立目录后的 Skill 产物，不需要安装 WPS 或连接外部账号。

## 项目结构

项目采用 Java 风格的 source set，同时保留标准 Python 包：

```text
src/
  main/
    python/wps_skills/
      cli/          # 命令解析、能力发现和构建入口
      client/       # 调用方的 Session Client
      core/         # 应用无关的契约和 Action Session 机制
      host/         # JSONL Session Host
      word/         # Word 契约、Adapter；windows/ 包拥有后端与会话组装
      excel/        # Excel 契约、Adapter；windows/ 包拥有后端与会话组装
      ppt/          # PPT 契约、Adapter；windows/ 包拥有后端与会话组装
      windows/      # 共享进程、文档协调、传输和桌面支持
    resources/
      wps_skills/word/windows/  # PowerShell bridge 和 Word 操作实现
      skills/wps-word/         # Word Skill 源文件、参考文档和薄入口
      skills/wps-excel/        # Excel Skill 源文件、参考文档和薄入口
      wps_skills/excel/windows/ # Excel 操作实现
      wps_skills/ppt/windows/   # PPT 原生操作实现
      skills/wps-ppt/          # PPT Skill 源文件、参考文档和薄入口
      wps_skills/windows/      # 共享桥接、原生文件身份和文档协调
  test/
    python/tests/   # 单元测试、真实验收和可执行 Python 测试辅助程序
    resources/      # PowerShell 测试资源与类型库能力证据
scripts/            # 仓库级薄入口
build/              # 生成的安装包和验收产物（忽略提交）
```

测试使用独立的 `tests` 命名空间，避免在测试发现时用另一个顶层 `wps_skills` 包遮蔽生产实现。

完整的目录约定、每个维护文件的作用与执行流程见 [文件结构说明](FILE_STRUCTURE.md)。

## Word 可见演示

在 Windows 桌面普通 PowerShell 的仓库根目录运行：

```powershell
$python = (Get-Command python.exe).Source
./scripts/demo/word.ps1 -PythonPath $python -Delay 1.5
```

演示依次写入标题、正文和原生表格，读回验证后显式保存，保留文档窗口。输出使用独立目录和唯一文件名；空白 DOCX 由演示夹具准备，全部内容通过正式 Action 写入。原生闭环验收使用 `python scripts/validate/word.py --output-dir build/word-acceptance`。

脚本已按 `build/`、`demo/`、`validate/` 分组，完整入口和参数见 [scripts/README.md](scripts/README.md)。

## Skill 与能力参考

- [Word Skill](src/main/resources/skills/wps-word/SKILL.md)：任务工作流程与能力发现。
- [会话使用说明](src/main/resources/skills/wps-word/references/session.md)：Python 客户端用法与可运行示例。
- [内容与格式](src/main/resources/skills/wps-word/references/content.md)：内容范围、revision、文字格式和单位。
- [验证与保存](src/main/resources/skills/wps-word/references/verification.md)：结果验证、持久化和失败处置。
- [Word 契约实现](src/main/python/wps_skills/word/contracts.py)：Action 的权威定义与正式能力集。
- [WPS Writer Type Library 快照](src/test/resources/wps_skills/word/type_library/wps_writer_api.py)：能力证据，不被 Runtime 导入，也不决定正式可用 Actions。

旧版组合式 Skill、全局 Manifest、多应用 Runtime、控制器、Linux/OpenXML 后端和旧验证脚本已完成切换。历史实现可从 Git 历史查看，当前系统不提供旧接口兼容层。
