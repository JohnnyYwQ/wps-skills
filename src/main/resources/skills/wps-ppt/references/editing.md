# PPT 编辑与验证

| 动作 | expectedToken 来源 | 核对结果 |
| --- | --- | --- |
| addSlide / moveSlide | listSlides | 新页或目标 SlideID 的最终位置及全部页顺序 |
| duplicateSlide / deleteSlide | 源页 getSlideInfo | 新页内容或源 ID 消失，剩余页顺序 |
| addTextBox / addShape | 目标页 getSlideInfo | 新增形状 ID、文本、位置、尺寸 |
| setShapeText / formatText / setShapeGeometry / deleteShape | 目标页 getSlideInfo | 对应 shapeId 的实际文本、字体、尺寸或删除事实 |

观察 token 只覆盖返回的字段。`listSlides` 覆盖页顺序和摘要；`getSlideInfo` 覆盖顶层形状、几何、层级、文本及汇总字体，不覆盖备注、动画、组内对象和逐字格式。用户编辑可能发生在观察之后；保存前仍需核对任务结果。删除页或形状应来自用户的明确删除意图，它会删除对象的全部内容。

结构操作最多 200 页；单页读取最多 100 个顶层形状，单个形状文本最多 10000 UTF-16 单元。超限明确失败，不返回截断的完整性声明。组、表格、图表的内部文字不是顶层文本框能力。

`setShapeText` 替换整个形状文字，换行使用 LF (`\n`)，可能重置富文本格式。`formatText` 对全部文字应用统一字体补丁，`latinName` 和 `eastAsianName` 分别设置西文和东亚字体；混合字体属性读取为 null，空文本不提供字体写后验证。

坐标和字号单位为 point。`setShapeGeometry` 保留 WPS 原生宽高比行为，因此不仅检查请求字段，还要检查其他尺寸是否符合任务。`addShape` 仅支持 rectangle、ellipse。颜色是 OLE RGB 整数：红 + 绿×256 + 蓝×65536，例如红色 255、蓝色 16711680。

每次修改返回实际读回结果；不是简单的“设置成功”。保存必须仍指向同一个 live Presentation 与已授权 backing file。外部 Save As 或无法证明的身份变化使会话终止。

填充/边框、段落/文本框、图片、表格、背景和备注的动作与 token 来源见 [common.md](common.md)。
