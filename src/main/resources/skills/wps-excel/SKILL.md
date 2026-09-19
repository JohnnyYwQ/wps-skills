---
name: wps-excel
description: 创建、读取、编辑和排版 Excel（.xlsx）工作簿，支持单元格数据、公式、区域格式、工作表及行列调整、工作簿保存和 PDF 导出。
---

# WPS Excel

根据用户目标和 Action 用途选择操作、安排顺序，再查询所选 schema 填写 JSON，提交后读取 Task Response，用自然语言报告结果。

下文 `<skill-dir>` 为本 Skill 目录，命令通过 Bash 调用 Python。

## 1. 根据用途选择 Actions 并安排顺序

从上下文确定新建意图或已有 `.xlsx` 的绝对路径，目标有歧义时才询问。用户未指定的常规表头、数字格式、字体和行列布局自主决定。

先根据下表选择完成用户目标所需的 Actions，确定步骤顺序及结果依赖。一个 Task 处理一个工作簿，分为获取文档、内容操作和交付三部分：

| 放置位置 | Action | 用途 |
| --- | --- | --- |
| document | `createWorkbook` | 新建空白工作簿。 |
| document | `openWorkbook` | 打开指定工作簿。 |
| steps | `getWorkbookInfo` | 读取工作簿状态和日期系统。 |
| steps | `listWorksheets` | 获取工作表名称与位置。 |
| steps | `getWorksheetInfo` | 读取工作表信息及结构修改 token。 |
| steps | `readRange` | 读取区域的值、公式、格式及区域 token。 |
| steps | `writeRange` | 写入矩形常量数据。 |
| steps | `setFormulas` | 写入矩形公式。 |
| steps | `calculateRange` | 计算指定区域。 |
| steps | `formatRange` | 设置数字、字体、颜色、对齐和换行。 |
| steps | `addWorksheet` / `renameWorksheet` / `copyWorksheet` / `moveWorksheet` / `deleteWorksheet` | 新增、重命名、复制、移动或删除工作表。 |
| steps | `insertRows` / `deleteRows` / `insertColumns` / `deleteColumns` | 插入或删除整行、整列。 |
| steps | `clearRange` | 清除值、格式或两者。 |
| steps | `mergeRange` / `unmergeRange` | 合并或取消合并区域。 |
| steps | `sortRange` / `filterRange` / `clearFilter` | 排序、筛选或清除筛选。 |
| steps | `setRowHeight` / `setColumnWidth` / `autoFitColumns` | 设置行高、列宽或自动列宽。 |
| steps | `findInRange` / `replaceInRange` | 查找文字或替换常量文字。 |
| steps | `copyRange` | 复制当前值到等大的不重叠区域，公式按计算结果复制。 |
| completion | `save` | 保存到已有工作簿的原路径。 |
| completion | `saveAs` | 首次保存或另存 XLSX。 |
| completion | `exportPdf` | 按当前打印设置导出 PDF。 |

用户未要求保存或导出时，使用 `completion: []`。新工作簿首次保存用 `saveAs`；PDF 导出不代表保存 XLSX。

读取用于理解内容或取得后续步骤所需的定位与 token；执行结果以 Task Response 为准，不额外安排验证性回读。

若必须先理解已有数据才能决定修改内容，先提交只读 Task，再根据结果编排修改 Task。新 Task 明确指定同一文件并重新取得工作表定位与 token。未保存的新工作簿没有跨 Task 接续入口，尽量一次编排完成。

## 2. 查询所选 Actions 的 schema

将已选出的 Action 名称传给查询脚本，批量取得填写参数和结果引用所需的定义：

```text
python "<skill-dir>/scripts/schema.py" createWorkbook listWorksheets readRange writeRange saveAs
```

输出 `actions` 中的 `parameters` 用于填写参数，`result` 描述 Action Response 的 `data`，`examples` 提供参数示例。顶层 `$defs` 是本次所需的公共定义；schema 中的 `{"$ref":"#/$defs/名称"}` 指向其中的定义。

输出被截断时，用同一脚本分批查询补齐，再填写相应参数。

## 3. 按 schema 填写请求

将已安排的步骤填入请求。顶层为 `app: "excel"`、`document`、`steps`、`completion`；仅在用户授权保存已有未保存修改时增加 `includeExistingChanges: true`。

| 部分 | 结构 |
| --- | --- |
| `document` | 一个 `createWorkbook` 或 `openWorkbook`；仅含 `address`、`params`。 |
| `steps` | 0–125 个有序步骤；每步仅含 `id`、`address`、`params`。 |
| `completion` | 0–2 个交付操作；每项仅含 `address`、`params`。`save`/`saveAs` 合计最多一个，`exportPdf` 最多一个；先保存后导出。 |

`address` 为 `{"app":"excel","action":"操作名称"}`，`params` 按查询结果填写，无参数也写 `{}`。不填 `version`、`taskId` 或 `checks`。步骤 id 使用 1–80 个 ASCII 字母、数字、下划线或连字符，以字母或数字开头；同一 Task 内唯一，避开 `task_document`、`task_save`、`task_pdf`。

### 内容与结果引用

- 工作表名称从 `listWorksheets` 获取，不猜测默认名称。已有工作表按用户指定或读取结果选择；新建工作簿的首张工作表可引用 `data.worksheets[0].name`。
- 区域使用工作表内的大写 A1 地址，如 `A1:C10`；单次最多 1000 个单元格，数据和公式矩阵的行列数必须与区域一致。
- `writeRange` 写常量，`setFormulas` 写公式；公式使用英文函数名和同工作表 A1 引用。日期按工作簿的 `date1904` 日期系统填写数字序列，并设置数字格式。
- 区域修改的 `expectedToken` 引用同一工作表、同一区域的最新 `readRange` 或修改结果。工作表及整行整列修改使用 `getWorksheetInfo` 的 token；`copyRange` 还需要目标区域的 token。
- token 只在本 Task 和对应观察范围未变化时有效。后续操作需要新 token 时引用同范围的修改结果，或重新读取；这是编排依赖，不是额外验证。
- 需要保存或导出的工作簿最多 20 张工作表，每张已用区域最多 1000 个单元格。

参数可引用本 Task 前序步骤的结果，**路径从 `data` 开始**，例如引用 read 步骤的 token：

```json
{"$ref":{"step":"read","path":["data","token"]}}
```

上述 Task 引用与 schema 的字符串 `$ref` 不同，不执行计算、循环或条件判断。数组路径使用数字下标，选取哪一项必须有依据；分批读取的部分结果不当作全部工作表或数据。

保存/导出参数按 schema 填写：输出使用绝对路径，`overwritePolicy` 使用 `failIfExists`，不覆盖已有目标。

下面是新建两行数据且不保存的完整请求：

```json
{
  "app": "excel",
  "document": {"address": {"app": "excel", "action": "createWorkbook"}, "params": {}},
  "steps": [
    {
      "id": "sheets",
      "address": {"app": "excel", "action": "listWorksheets"},
      "params": {"offset": 0, "limit": 100}
    },
    {
      "id": "read",
      "address": {"app": "excel", "action": "readRange"},
      "params": {
        "sheet": {"$ref": {"step": "sheets", "path": ["data", "worksheets", 0, "name"]}},
        "address": "A1:B2"
      }
    },
    {
      "id": "write",
      "address": {"app": "excel", "action": "writeRange"},
      "params": {
        "sheet": {"$ref": {"step": "sheets", "path": ["data", "worksheets", 0, "name"]}},
        "address": "A1:B2",
        "expectedToken": {"$ref": {"step": "read", "path": ["data", "token"]}},
        "values": [["项目", "数量"], ["星河知识库", 3]]
      }
    }
  ],
  "completion": []
}
```

## 4. 提交并读取结果

用宿主文件写入工具，将完整 JSON 写入独立 UTF-8 文件。每份请求使用新路径，已提交路径不修改、不复用。

```text
python "<skill-dir>/scripts/excel.py" --app excel --task-file "<本次工作目录>/request.json"
```

读取 Task Response：

| 字段 | 用途 |
| --- | --- |
| `outcome`、`state` | 总体执行结果和状态。 |
| `document`、`steps[]`、`completion.save`、`completion.pdf` | 各操作的 `state` 和 `response`；交付项为 `null` 表示未请求。 |
| `stop`、失败操作的 `response.error` | 停止位置和错误原因。 |
| 保存/导出操作的 `response.data.artifact.path` | 实际交付路径。 |

**提交后不自动重试。** 根据 Task Response 报告已完成、未完成的内容和原因；请求被拒绝、执行失败或结果不确定时停止。只有用户明确要求继续或调整，才开始新的请求。为理解已有数据而安排的只读 Task 成功后继续修改，不属于失败重试。

例如，Task Response 中保存或导出操作报告目标父目录不存在时，据此报告已完成、未完成的内容和原因并停止，不擅自创建目录、更换目标路径或重新提交 Task。

没有完整响应或命令超时时，用原输入路径查询回执；即使输入文件已不存在，也使用原路径：

```text
python "<skill-dir>/scripts/excel.py" --app excel --task-status-file "<本次工作目录>/request.json"
```

仍未取得完整 Task Response 时，报告无法确认并停止。

## 5. 自然语言交付

简要说明完成了哪些、哪些未完成，以及实际产物路径和保存状态。`not_executed` 是未执行，`unknown` 是无法确认；失败不表示已回滚。以回执为依据，不复述 JSON 或内部实现过程。
