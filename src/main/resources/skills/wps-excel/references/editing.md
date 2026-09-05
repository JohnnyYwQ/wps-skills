# 常用区域编辑

所有区域 action 使用准确 `sheet`、`address`；修改还需要同一区域 `readRange` 返回的 `expectedToken`。每次最多 1000 个单元格。返回的是 WPS 实际回读结果。

| Action | 参数及边界 |
| --- | --- |
| `clearRange` | `mode`: contents / formats / all |
| `copyRange` | `targetSheet`、`targetAddress`、`targetToken`；源/目标等大且不重叠，分别读 token；复制计算后的字面值，不复制公式/样式；拒绝错误值 |
| `findInRange` | `text`、`matchCase`、`wholeCell`、`lookIn`: values / formulas；字面匹配，返回单元格地址和文本；只读不需要 expectedToken |
| `replaceInRange` | `text`、`replacement`、`matchCase`、`wholeCell`；只替换字符串常量，不改公式和数字；不把输入当正则或通配符 |
| `sortRange` | `column` 为区域内 1 起始列号，`order`: ascending / descending，`header` 为布尔值；整行一起排序，支持常量，拒绝公式和错误值；文本顺序由 WPS 区域设置决定 |
| `filterRange` | `column`、字符串 `value`；首行为标题，按字面等于筛选；`*` / `?` 为普通字符；已有筛选必须属于同一区域，其他列的条件仍生效 |
| `clearFilter` | 清除这个区域对应的工作表筛选；拒绝操作其他区域的筛选 |
| `mergeRange` | 仅左上角可以非空，防止合并丢数据 |
| `unmergeRange` | 必须包含完整合并区域，不自动扩展选择 |
| `setRowHeight` | `height`，单位磅，影响区域经过的整行 |
| `setColumnWidth` | `width`，字符宽度单位，影响区域经过的整列 |
| `autoFitColumns` | 按指定区域的内容自动适配列宽 |

尺寸受 WPS 像素取整影响，检查返回的实际 `rowHeight` / `columnWidth`；允许小幅量化误差。筛选后用 `rowHidden` 检查显示行。合并与结构修改后重新读取。响应为 unknown 时先只读核验，不重放修改。
