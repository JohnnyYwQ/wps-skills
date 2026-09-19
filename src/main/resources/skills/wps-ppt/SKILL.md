---
name: wps-ppt
description: 创建、读取、编辑和排版 PowerPoint（.pptx）演示文稿，支持幻灯片、文字与形状格式、图片和表格、讲者备注、文稿保存及 PDF/PNG 导出。
---

# WPS PPT

根据用户目标和 Action 用途选择操作、安排顺序，再查询所选 schema 填写 JSON，提交后读取 Task Response，用自然语言报告结果。

下文 `<skill-dir>` 为本 Skill 目录，命令通过 Bash 调用 Python。

## 1. 根据用途选择 Actions 并安排顺序

从上下文确定新建意图或已有 `.pptx` 的绝对路径，目标有歧义时才询问。用户未指定的常规页面安排、字体、配色和排版自主决定。

先根据下表选择完成用户目标所需的 Actions，确定步骤顺序及结果依赖。一个 Task 处理一份演示文稿，分为获取文档、内容操作和交付三部分：

| 放置位置 | Action | 用途 |
| --- | --- | --- |
| document | `createPresentation` | 新建空白演示文稿。 |
| document | `openPresentation` | 打开指定演示文稿。 |
| steps | `getPresentationInfo` | 读取文稿状态和页面尺寸。 |
| steps | `listSlides` | 读取幻灯片 ID、顺序与列表 token。 |
| steps | `getSlideInfo` | 读取幻灯片形状及 slide token。 |
| steps | `addSlide` / `duplicateSlide` / `moveSlide` / `deleteSlide` | 新增、复制、移动或删除幻灯片。 |
| steps | `addTextBox` / `addShape` / `addImage` | 新增文本框、几何形状或图片。 |
| steps | `setShapeText` / `replaceText` / `findText` | 设置、替换或查找文字。 |
| steps | `formatText` | 设置中西文字体、字号和字符样式。 |
| steps | `setShapeGeometry` / `deleteShape` / `renameShape` | 调整位置尺寸、删除或重命名形状。 |
| steps | `getShapeStyle` | 读取形状外观、段落、文本框布局和 style token。 |
| steps | `formatShape` / `formatParagraph` / `setTextBoxLayout` | 设置形状外观、段落或文本框布局。 |
| steps | `setShapeOrder` / `alignShapes` / `distributeShapes` | 调整层次、对齐或分布形状。 |
| steps | `getSlideSettings` / `setSlideSettings` | 读取或修改幻灯片名称、背景和放映可见性。 |
| steps | `getSlideNotes` / `setSlideNotes` | 读取或修改讲者备注。 |
| steps | `addTable` / `readTable` / `writeTable` | 插入、读取或修改表格。 |
| steps | `exportSlideImage` | 在 steps 中导出指定幻灯片 PNG。 |
| completion | `save` | 保存到已有文稿的原路径。 |
| completion | `saveAs` | 首次保存或另存 PPTX。 |
| completion | `exportPdf` | 导出整份 PDF。 |

用户未要求保存或导出时，使用 `completion: []`。新文稿首次保存用 `saveAs`；PDF 导出不代表保存 PPTX。

读取用于理解内容或取得后续步骤所需的定位与 token；执行结果以 Task Response 为准，不额外安排验证性回读。

若必须先理解已有页面才能决定修改内容，先提交只读 Task，再根据结果编排修改 Task。新 Task 明确指定同一文件并重新取得幻灯片、形状定位与 token。未保存的新文稿没有跨 Task 接续入口，尽量一次编排完成。

## 2. 查询所选 Actions 的 schema

将已选出的 Action 名称传给查询脚本，批量取得填写参数和结果引用所需的定义：

```text
python "<skill-dir>/scripts/schema.py" createPresentation listSlides addSlide getSlideInfo addTextBox saveAs
```

输出 `actions` 中的 `parameters` 用于填写参数，`result` 描述 Action Response 的 `data`，`examples` 提供参数示例。顶层 `$defs` 是本次所需的公共定义；schema 中的 `{"$ref":"#/$defs/名称"}` 指向其中的定义。

输出被截断时，用同一脚本分批查询补齐，再填写相应参数。

## 3. 按 schema 填写请求

将已安排的步骤填入请求。顶层为 `app: "ppt"`、`document`、`steps`、`completion`；仅在用户授权保存已有未保存修改时增加 `includeExistingChanges: true`。

| 部分 | 结构 |
| --- | --- |
| `document` | 一个 `createPresentation` 或 `openPresentation`；仅含 `address`、`params`。 |
| `steps` | 0–125 个有序步骤；每步仅含 `id`、`address`、`params`。 |
| `completion` | 0–2 个交付操作；每项仅含 `address`、`params`。`save`/`saveAs` 合计最多一个，`exportPdf` 最多一个；先保存后导出。 |

`address` 为 `{"app":"ppt","action":"操作名称"}`，`params` 按查询结果填写，无参数也写 `{}`。不填 `version`、`taskId` 或 `checks`。步骤 id 使用 1–80 个 ASCII 字母、数字、下划线或连字符，以字母或数字开头；同一 Task 内唯一，避开 `task_document`、`task_save`、`task_pdf`。

### 内容与结果引用

- `slideId`、`shapeId` 使用前序结果中的实际 ID：例如 `listSlides` 的 `data.slides[].id`、`getSlideInfo` 的 `data.shapes[].id`。形状 ID 只在所属幻灯片内使用；ID 不等于数组下标，插入和移动的 `position` 从 1 开始。
- 坐标、尺寸和字号使用 pt；中西文字体分别填写。新增对象及设置文字后，再安排所需的字符格式、段落和布局操作。
- `addSlide`、`moveSlide` 使用 `listSlides` 的 token；`duplicateSlide`、`deleteSlide` 及形状新增、文字、字符格式、几何、顺序操作使用 `getSlideInfo` 的 token。
- `formatShape`、`formatParagraph`、`setTextBoxLayout` 使用 `getShapeStyle` 的 token；`setSlideSettings`、`setSlideNotes`、`writeTable` 分别使用 `getSlideSettings`、`getSlideNotes`、`readTable` 的 token。
- token 只在本 Task 和对应观察范围未变化时有效。同范围修改结果的新 token 可以继续引用，否则重新读取；列表、幻灯片、样式、备注和表格的 token 不混用。
- 最多 200 页，每页最多 100 个顶层形状；表格最多 20 行、10 列、总计 100 个单元格。具体对象类型与可设置属性按所选 schema 填写。
- `exportSlideImage` 放在 `steps`，引用目标幻灯片 ID；PNG 导出不代表保存 PPTX。

参数可引用本 Task 前序步骤的结果，**路径从 `data` 开始**，例如引用 read 步骤的 token：

```json
{"$ref":{"step":"read","path":["data","token"]}}
```

上述 Task 引用与 schema 的字符串 `$ref` 不同，不执行计算、循环或条件判断。数组路径使用数字下标，选取哪一项必须有依据，不猜测新建对象的 ID。

保存/导出参数按 schema 填写：输出使用绝对路径，`overwritePolicy` 使用 `failIfExists`，不覆盖已有目标。

下面是新建一页、插入标题且不保存的完整请求：

```json
{
  "app": "ppt",
  "document": {"address": {"app": "ppt", "action": "createPresentation"}, "params": {}},
  "steps": [
    {
      "id": "slides",
      "address": {"app": "ppt", "action": "listSlides"},
      "params": {}
    },
    {
      "id": "add",
      "address": {"app": "ppt", "action": "addSlide"},
      "params": {
        "position": 1,
        "expectedToken": {"$ref": {"step": "slides", "path": ["data", "token"]}}
      }
    },
    {
      "id": "read",
      "address": {"app": "ppt", "action": "getSlideInfo"},
      "params": {"slideId": {"$ref": {"step": "add", "path": ["data", "slides", 0, "id"]}}}
    },
    {
      "id": "title",
      "address": {"app": "ppt", "action": "addTextBox"},
      "params": {
        "slideId": {"$ref": {"step": "add", "path": ["data", "slides", 0, "id"]}},
        "expectedToken": {"$ref": {"step": "read", "path": ["data", "token"]}},
        "text": "星河知识库项目启动",
        "left": 40, "top": 40, "width": 600, "height": 80
      }
    }
  ],
  "completion": []
}
```

## 4. 提交并读取结果

用宿主文件写入工具，将完整 JSON 写入独立 UTF-8 文件。每份请求使用新路径，已提交路径不修改、不复用。

```text
python "<skill-dir>/scripts/ppt.py" --app ppt --task-file "<本次工作目录>/request.json"
```

读取 Task Response：

| 字段 | 用途 |
| --- | --- |
| `outcome`、`state` | 总体执行结果和状态。 |
| `document`、`steps[]`、`completion.save`、`completion.pdf` | 各操作的 `state` 和 `response`；交付项为 `null` 表示未请求。 |
| `stop`、失败操作的 `response.error` | 停止位置和错误原因。 |
| 保存/导出操作的 `response.data.artifact.path` | 实际交付路径。 |

**提交后不自动重试。** 根据 Task Response 报告已完成、未完成的内容和原因；请求被拒绝、执行失败或结果不确定时停止。只有用户明确要求继续或调整，才开始新的请求。为理解已有页面而安排的只读 Task 成功后继续修改，不属于失败重试。

例如，Task Response 中保存或导出操作报告目标父目录不存在时，据此报告已完成、未完成的内容和原因并停止，不擅自创建目录、更换目标路径或重新提交 Task。

没有完整响应或命令超时时，用原输入路径查询回执；即使输入文件已不存在，也使用原路径：

```text
python "<skill-dir>/scripts/ppt.py" --app ppt --task-status-file "<本次工作目录>/request.json"
```

仍未取得完整 Task Response 时，报告无法确认并停止。

## 5. 自然语言交付

简要说明完成了哪些、哪些未完成，以及实际产物路径和保存状态。`not_executed` 是未执行，`unknown` 是无法确认；失败不表示已回滚。以回执为依据，不复述 JSON 或内部实现过程。
