# 常用排版、图片、表格和备注

先批量解析任务需要的完整 Action Contracts。不同属性使用不同观察 token，不能互换：

| 编辑任务 | 先读取 | 动作及验证 |
| --- | --- | --- |
| 填充与边框 | getShapeStyle | formatShape；检查颜色、透明度、可见性和线宽 |
| 段落与文本框 | getShapeStyle | formatParagraph / setTextBoxLayout；检查每个段落及边距、换行、垂直对齐 |
| 形状改名、层级、对齐、分布 | getSlideInfo | renameShape / setShapeOrder / alignShapes / distributeShapes；检查稳定 ID、zOrder 和边界 |
| 背景、隐藏和页名称 | getSlideSettings | setSlideSettings；检查背景继承和隐藏状态 |
| 演讲备注 | getSlideNotes | setSlideNotes；检查完整备注正文 |
| 查找替换 | getSlideInfo 或 findText | replaceText；检查替换次数和完整结果文字 |
| 插入图片、表格 | getSlideInfo | addImage / addTable；检查新增对象和内容 |
| 修改已有表格 | readTable | writeTable；逐格检查整个矩阵 |

## 形状与文本排版

`getShapeStyle` 的 token 覆盖该形状的摘要、填充/边框、文本框和各段落属性。它支持至多 100 个段落，不读取组内部或表格内部格式。`fillColor` 会将填充改成纯色，替换已有渐变或图片填充；颜色使用 OLE RGB 整数。设置填充颜色/透明度会启用填充，设置边框颜色会启用边框；不能同时要求对应可见性为 false。隐藏或非纯色填充的颜色、隐藏边框颜色读取为 null。透明度范围为 0–1，线宽、边距和段前段后间距单位为 point。

`formatParagraph` 作用于全部段落，设置段前/段后间距时明确使用 point。项目符号仅控制显示与否，具体样式由 WPS 的原生行为决定。`setTextBoxLayout` 保留 AutoSize 和原生尺寸行为；排版后也应检查形状尺寸。

`setShapeOrder` 支持 front/back，`zOrder=1` 是最底层。`alignShapes` 按指定形状集合的包围框对齐，支持左/中/右/上/中部/下；不使用当前 UI 选区。`distributeShapes` 保持外侧形状，均匀分配边缘间距。只支持未旋转、未分组的顶层形状，最多 20 个；分布至少需要 3 个形状，空间不足以形成非负间距时拒绝。

## 背景和备注

`setSlideSettings` 的 `hidden` 控制放映时是否跳过该页，编辑窗口仍然可见。指定 `backgroundColor` 会关闭母版背景继承并设置纯色；不能同时要求继承母版背景。

备注仅操作 NotesPage 中唯一的正文占位符，不改日期、页脚或幻灯片图像占位符。`setSlideNotes` 替换完整正文，可能重置原有备注富文本格式。读取和写入最多 10000 UTF-16 单元。

## 查找替换

`findText` 在一页顶层文本形状中做区分大小写的字面匹配，最多 100 处。返回零起始的 UTF-16 偏移，不是 Python 字符索引，也不是文稿地址。它不查组、表格、备注或其他页。

`replaceText` 在指定 `shapeId` 内替换所有非重叠匹配；不接受任意 COM 表达式或正则。原生范围编辑从后向前执行，保留替换范围外文字；替换文字的字体继承按 WPS 行为处理，必要时随后显式设置格式。若原生换行不能对应到规范化偏移，动作在写入前拒绝。结果包含替换次数和完整页读回内容。

## 图片与表格

`addImage` 接收 Windows 主机上已有 PNG/JPEG 的绝对路径，图片最多 20 MiB、4000 万像素，必须嵌入文稿而非外部链接。明确指定位置和宽高；宽高比例不一致会拉伸图片。插入时验证实际图片内容，保存后检查嵌入结果。

`addTable` / `readTable` / `writeTable` 处理原生 PPT 表格的文字矩阵。上限为 20 行、10 列、100 个单元格，每格 2000 UTF-16 单元、整表 20000 单元。数值也作为文字提供，不引入 Excel 公式或工作簿。

`writeTable` 必须给出与已有行列数完全一致的矩阵，它替换每格全部文字，可能重置单元格富文本格式。token 覆盖表格 ID、行列数和文字，不覆盖样式或合并单元格几何。合并单元格可能使写入失败或只完成部分修改；出现 unknown 后先只读检查，禁止自动重试。
