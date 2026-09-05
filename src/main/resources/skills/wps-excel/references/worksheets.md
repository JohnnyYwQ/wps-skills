# 工作表和结构编辑

`getWorksheetInfo({sheet})` 返回准确名称、1 起始索引、可见状态、已用区域地址和 `token`。该 token 观察工作簿的工作表名称/顺序、保护状态以及目标表已用区域的值、公式、格式与行列尺寸。读取和结构操作的已用区域上限均为 1000 个单元格；超限返回 `RANGE_UNSUPPORTED`，不能用局部 token 绕过整表检查。

下列 action 均传 `sheet` 和刚读取的 `expectedToken`：

| Action | 额外参数与行为 |
| --- | --- |
| `addWorksheet` | `name`；在指定表后创建空表，名称不得重复 |
| `renameWorksheet` | `name`；修改确切工作表名称 |
| `copyWorksheet` | `name`；在原表后复制，在同一工作簿内命名 |
| `moveWorksheet` | `index`；移动到指定的 1 起始位置 |
| `deleteWorksheet` | 删除已观察的工作表，拒绝删除最后一张可见表 |
| `insertRows` / `deleteRows` | `start`、`count`；整行插入/删除 |
| `insertColumns` / `deleteColumns` | `start`、`count`；整列插入/删除，列号是 1 起始整数 |

行列变更最多一次 100 行/列，并检查变更后仍可完整观察。合并和数组公式区域不支持结构修改。操作影响整行/列，WPS 可能调整其他公式的引用；完成后重新读取依赖区域验证用户结果。

结构操作成功后清除旧工作表缓存和 token。除删除返回 `deleted` 外，均返回变更后的工作表信息。删除、改名、移动前重新读取；不自动保存，任务最后显式 `save`。
