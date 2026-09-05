# 区域、数据和格式

`sheet` 是准确的工作表名称，`address` 是大写、不带 `$` 的工作表内 A1 单元格或矩形，例如 `A1`、`A1:C20`。不接受整行、整列、命名范围、跨表地址或不连续区域。每次最多 1000 个单元格；大的任务分区处理，每区读取并验证后再继续。

`readRange` 返回 `cells` 二维数组，每个单元格包括：

- `value`：null、字符串、布尔值或数字；日期是序列数字，结合 `getWorkbookInfo.date1904` 和 `numberFormat` 解读。
- `formula`：公式字符串或 null。写入文本以 `writeRange` 为准，不自动解释为公式。
- `text`：WPS 显示文本，可能受列宽影响，不能用于精确数值验证。
- `errorCode`：COM 错误代码或 null；错误单元格的 value 为 null。
- `numberFormat`、`bold`、`italic`、`fontSize`、`fontColor`、`fillColor`、`wrapText`、水平/垂直对齐：实际格式。
- `merged`、`rowHidden`、`columnHidden`、`rowHeight`、`columnWidth`：合并、隐藏和尺寸状态。

`token` 是对本 Session、工作表对象和这一矩形观察结果的校验标记，不是整个工作簿的 Content Revision，也不表示文档身份。所有区域修改 action 都要求同一区域最新读取的 `expectedToken`。其他区域变化不一定使它失效；不能据此断言整个工作簿没有变化。工作表结构修改后所有旧 token 失效，重新发现并读取。

`writeRange.values` 必须与矩形行列数完全一致，不能靠自动扩展或截断补齐。null 清除内容，字符串保持文本类型（包括以 `=` 开头的文本）。单个文本值最多 4096 个字符；矩阵总文本大小受契约限制。

`formatRange.format` 接受非空补丁：`numberFormat`、`bold`、`italic`、`fontSize`、`fontColor`、`fillColor`、`wrapText`、`horizontalAlignment`（general/left/center/right）、`verticalAlignment`（top/center/bottom）。颜色使用 OLE 整数 `red + 256*green + 65536*blue`，例如红色 255。读到其他对齐方式返回 `other`，不能用 `other` 写入。

拒绝向只读工作簿、受保护工作表和数组公式区域写入。合并/取消合并由专门 action 处理；普通写入仍拒绝合并区域。动态数组兼容性需要按安装的 WPS 版本验证，不能用拆分写入绕过拒绝。
