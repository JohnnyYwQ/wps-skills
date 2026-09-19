---
name: wps-word
description: 创建、读取、编辑和排版 Word（.docx）文档，支持文字与页面格式调整、表格和图片插入、文档保存及 PDF 导出。
---

# WPS Word

根据用户目标和 Action 用途选择操作、安排顺序，再查询所选 schema 填写 JSON，提交后读取 Task Response，用自然语言报告结果。

下文 `<skill-dir>` 为本 Skill 目录，命令通过 Bash 调用 Python。

## 1. 根据用途选择 Actions 并安排顺序

从上下文确定新建意图或已有 `.docx` 的绝对路径，目标有歧义时才询问。用户未指定的常规字体、标题层级和排版自主决定。

先根据下表选择完成用户目标所需的 Actions，确定步骤顺序及结果依赖。一个 Task 处理一个文档，分为获取文档、内容操作和交付三部分：

| 放置位置 | Action | 用途 |
| --- | --- | --- |
| document | `createDocument` | 新建空白文档。 |
| document | `openDocument` | 打开指定的已有文档。 |
| steps | `writeContent` | 插入文字、段落或标题，并设置文字与段落格式。 |
| steps | `inspectDocument` | 读取正文、格式、结构、页面布局和文档状态。 |
| steps | `findContent` | 按文字查找匹配内容，取得范围。 |
| steps | `replaceContent` | 按文本匹配或指定范围替换、删除内容；结构化段落替换支持段落格式。 |
| steps | `insertTable` | 插入矩形文字表格。 |
| steps | `insertImage` | 插入 PNG/JPEG，设置尺寸、放置方式和替代文本。 |
| steps | `setHeaderFooter` | 设置所选节的页眉和页脚。 |
| steps | `setPageLayout` | 设置所选节的方向和页边距。 |
| steps | `insertBreak` | 插入分页符或分节符。 |
| completion | `save` | 保存到已有文档的原路径。 |
| completion | `saveAs` | 首次保存或另存 DOCX。 |
| completion | `exportPdf` | 导出整篇 PDF。 |

用户未要求保存或导出时，使用 `completion: []`。新文档首次保存用 `saveAs`；PDF 导出不代表保存 DOCX。

读取用于理解内容或取得后续步骤所需的范围；执行结果以 Task Response 为准，不额外安排验证性回读。

若必须先理解原文才能决定修改内容，先提交只读 Task，再根据结果编排修改 Task。新 Task 明确指定同一文件并重新取得范围。未保存的新文档没有跨 Task 接续入口，尽量一次编排完成。

## 2. 查询所选 Actions 的 schema

将已选出的 Action 名称传给查询脚本，批量取得填写参数和结果引用所需的定义：

```text
python "<skill-dir>/scripts/schema.py" createDocument writeContent saveAs
```

输出 `actions` 中的 `parameters` 用于填写参数，`result` 描述 Action Response 的 `data`，`examples` 提供参数示例。顶层 `$defs` 是本次所需的公共定义；schema 中的 `{"$ref":"#/$defs/名称"}` 指向其中的定义。

输出被截断时，用同一脚本分批查询补齐，再填写相应参数。

## 3. 按 schema 填写请求

将已安排的步骤填入请求。顶层为 `app: "word"`、`document`、`steps`、`completion`；仅在用户授权保存已有未保存修改时增加 `includeExistingChanges: true`。

| 部分 | 结构 |
| --- | --- |
| `document` | 一个 `createDocument` 或 `openDocument`；仅含 `address`、`params`。 |
| `steps` | 0–125 个有序步骤；每步仅含 `id`、`address`、`params`。 |
| `completion` | 0–2 个交付操作；每项仅含 `address`、`params`。`save`/`saveAs` 合计最多一个，`exportPdf` 最多一个；先保存后导出。 |

`address` 为 `{"app":"word","action":"操作名称"}`，`params` 按查询结果填写，无参数也写 `{}`。不填 `version`、`taskId` 或 `checks`。步骤 id 使用 1–80 个 ASCII 字母、数字、下划线或连字符，以字母或数字开头；同一 Task 内唯一，避开 `task_document`、`task_save`、`task_pdf`。

### 内容与结果引用

- 字符格式放在 run 的 `format`，段落格式放在 paragraph/heading 的 `format`；标题使用 `heading` 和 `level`。
- 参数可引用本 Task 前序步骤的结果。**路径从 `data` 开始**，例如引用 inspect 步骤的第一段范围：

```json
{"$ref":{"step":"inspect","path":["data","paragraphs",0,"range"]}}
```

- 上述 Task 引用取前序结果值，与 schema 的字符串 `$ref` 不同；它不执行计算、循环或内容判断。匹配数量和数组索引必须有依据。
- 范围使用读取或查找结果，不从正文字符串手算坐标。范围只在生成它的 Task 和有效 revision 内使用；内容修改后需要新范围时重新定位。
- 读取结果截断时不当作全文。`remainingRange` 也受相同的 Task/revision 约束，不能直接带到下一 Task。

保存/导出参数按 schema 填写：输出使用绝对路径；允许重名换名选 `renameIfExists`，必须保留指定名称且不能覆盖选 `failIfExists`，覆盖 PDF 的 `replaceExisting` 仅在用户明确授权时使用。

下面是新建文字且不保存的完整请求：

```json
{
  "app": "word",
  "document": {"address": {"app": "word", "action": "createDocument"}, "params": {}},
  "steps": [
    {
      "id": "write",
      "address": {"app": "word", "action": "writeContent"},
      "params": {
        "anchor": {"kind": "documentEnd"},
        "blocks": [{"kind": "paragraph", "runs": [{"text": "项目周报"}]}]
      }
    }
  ],
  "completion": []
}
```

## 4. 提交并读取结果

用宿主文件写入工具，将完整 JSON 写入独立 UTF-8 文件。每份请求使用新路径，已提交路径不修改、不复用。

```text
python "<skill-dir>/scripts/word.py" --app word --task-file "<本次工作目录>/request.json"
```

读取 Task Response：

| 字段 | 用途 |
| --- | --- |
| `outcome`、`state` | 总体执行结果和状态。 |
| `document`、`steps[]`、`completion.save`、`completion.pdf` | 各操作的 `state` 和 `response`；交付项为 `null` 表示未请求。 |
| `stop`、失败操作的 `response.error` | 停止位置和错误原因。 |
| 保存/导出操作的 `response.data.artifact.path` | 实际交付路径。 |

**提交后不自动重试。** 根据 Task Response 报告已完成、未完成的内容和原因；请求被拒绝、执行失败或结果不确定时停止。只有用户明确要求继续或调整，才开始新的请求。为理解原文而安排的只读 Task 成功后继续修改，不属于失败重试。

例如，Task Response 中保存或导出操作报告目标父目录不存在时，据此报告已完成、未完成的内容和原因并停止，不擅自创建目录、更换目标路径或重新提交 Task。

没有完整响应或命令超时时，用原输入路径查询回执；即使输入文件已不存在，也使用原路径：

```text
python "<skill-dir>/scripts/word.py" --app word --task-status-file "<本次工作目录>/request.json"
```

仍未取得完整 Task Response 时，报告无法确认并停止。

## 5. 自然语言交付

简要说明完成了哪些、哪些未完成，以及实际产物路径和保存状态。`not_executed` 是未执行，`unknown` 是无法确认；失败不表示已回滚。以回执为依据，不复述 JSON 或内部实现过程。
