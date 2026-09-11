---
name: wps-excel
description: 通过 Windows WPS 表格新建或编辑 Excel 工作簿，管理工作表、数据、公式和格式，保存、另存为及导出 PDF。用户要求在 WPS 中制作或操作 .xlsx 文件时使用。
---

# WPS Excel

使用本 Skill 的 `scripts/excel.py` 和 Excel Application Contract Set。路径相对于本文件；先定位 Skill 的绝对目录。执行需要 Windows、Python 3.8+、原生 Windows PowerShell 和注册为 `KET.Application` 的 WPS 表格；不需要 pywin32。安装时保留完整 `runtime/`。能力查询可在 macOS/Linux 执行。

## 确定目标与交付

已有文件使用 Windows 主机上 `.xlsx` 的绝对路径调用 `openWorkbook`；明确新建时使用 `createWorkbook`，得到一个含一张空白工作表的未保存工作簿。一个 Session 绑定一个确切工作簿，不猜测活动窗口，不把打开失败当成新建。当前不提供图表、透视表或 `.xls` / `.xlsm` / `.et` 编辑。

修改已有文件默认需要显式 `save`，除非用户要求保留未保存状态。目标已有未保存修改且本次需要保存时，在第一次修改前说明这些修改也会一起保存并取得确认；沿用已有明确授权。只读任务不保存。

首次保存或另存为前读 [references/persistence.md](references/persistence.md)。`saveAs` 只写入尚不存在的 `.xlsx`，之后 `save` 保存当前路径。导出不会替代源文件保存；用户未提供新建文件的输出路径时，可按意图保留未保存状态并如实报告。

## 发现能力

```powershell
python "<skill-dir>/scripts/excel.py" --app excel --index
python "<skill-dir>/scripts/excel.py" --app excel --resolve openWorkbook listWorksheets readRange writeRange save
```

查询不会启动 WPS。以返回的完整契约为准，只有所需契约全部解析成功后才执行。若返回 `WPS_DISCOVERY_UNAVAILABLE`，该构建尚未通过真机准入，不调用候选契约、不绕过桥接直接使用 COM。源码类型库只是能力证据。

## 使用一个会话

执行前读 [references/session.md](references/session.md)。直接调用包内 `scripts/excel.py` 的 `--start`、`--call`、`--status`、`--close`；Agent 只提供 Action 和 JSON 参数，无需编写任务脚本。一个任务复用返回的 `handle`，每次读取完整响应后才使用 `nextStep` 决定下一步。

先用 `getWorkbookInfo` 和 `listWorksheets` 了解工作簿；工作表名称必须使用返回的准确名称。区域操作前读 [references/ranges.md](references/ranges.md)，先 `readRange`，再把返回的 `token` 作为同一区域修改的 `expectedToken`。使用明确的工作表名称和 A1 矩形；不使用当前选区。

工作表增删、复制、改名、移动及行列插入删除前，读 [references/worksheets.md](references/worksheets.md)，使用 `getWorksheetInfo` 的整表观察 token。查找替换、复制、合并、排序筛选和尺寸调整读 [references/editing.md](references/editing.md)。

需要公式时读 [references/formulas.md](references/formulas.md)。修改后检查返回的实际值、公式和格式。保存前按 [references/verification.md](references/verification.md) 验证用户要求；显式保存后检查 artifact 和保存状态。

## 结束与异常

结束 Session 仅释放资源，不保存或关闭工作簿。报告修改、验证、保存路径和仍然打开的状态。Session 清理结果与用户文档任务结果分别判断。

命令 JSON 中的 `response.error` 保留实际失败结果。`unknown` 或丢失响应表示可能部分写入，不自动重试；会话可用时先只读检查。`STALE_RANGE` 需要重新读取、重新判断用户目标后再决定修改。文档关闭、桥接损坏或绑定身份变化会结束 Session，不跟随新路径、不重建丢失内容、不绕过 Lease 或 Quarantine。

运行日志位置和自定义方法见 [references/logging.md](references/logging.md)。
