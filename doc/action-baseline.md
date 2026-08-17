# WPS Action 基线

Issue #3 冻结了迁移开始时的能力事实。机器可读文件是唯一依据；本文只解释边界和审计结果，不重复维护 Action 列表。

## 事实来源

- `skills/wps-office/references/action-manifest.json`：公开 WPS Action 契约。每项记录应用归属、参数 Schema、结果 Schema、前置条件和 `read` / `write` / `destructive` 风险等级。
- `doc/migration/legacy-tool-action-map.json`：一次性迁移对照。它记录全部旧 WPS Tool 到原始 WPS Action 名称的映射，以及当前 Platform Bridge 的已解释差异。
- `scripts/validate_action_manifest.py`：Python 3.9+ 标准库校验入口。它直接扫描旧 TypeScript、JavaScript Add-in 和 PowerShell 路径，避免手工维护另一份能力计数。

公开接口只使用 manifest 中的 `action`，例如 `getCellValue`。旧 WPS Tool 名称只存在于迁移对照中。

## 冻结时的审计结果

| 范围 | 数量 | 说明 |
|---|---:|---|
| WPS Action manifest | 251 | 241 个 PowerShell dispatch，加 10 个旧 Tool 已调用但两个 Bridge 均未实现的 Action |
| 旧 WPS Tool | 250 | 241 个直接映射、6 个不产生 WPS Action 的本地能力、3 个已解释名称冲突 |
| JavaScript Add-in dispatch | 227 | 所有名称都进入 manifest |
| PowerShell 唯一 dispatch | 241 | 源文件有 245 个分支，4 个 Action 各重复一次 |
| JavaScript 已解释缺口 | 24 | 14 个 PowerShell-only Action，以及 10 个尚无 Bridge 实现的 Tool Action |
| PowerShell 已解释缺口 | 10 | 旧 Tool 已调用、但尚无 Bridge 实现的 Action |

原迁移分析中的“JavaScript 比 PowerShell 少 14 个”仍然成立。额外的 10 个双 Bridge 缺口来自 Tool handler：`autoSum`、`evaluateFormula`、`insertSectionBreak`、`setFontColor`、`setLineSpacing`、`setShapeFill`、`setSlideSize`、`setSlideTheme`、`setTextColor`、`setZoom`。

PowerShell 重复分支是 `create3DText`、`set3DDepth`、`set3DMaterial`、`set3DRotation`。manifest 对每个名称只保留一个规范 Action；迁移加载项时必须合并并验证重复实现，不能把重复名称公开成两个接口。

两个 Bridge 对 `addArrow` 使用同一名称却保留了不同语义：JavaScript 以起止坐标绘制线箭头，PowerShell 以边界框插入块箭头。manifest 暂时接受两组参数，迁移对照把该差异记录为 `contract_conflicts`；统一 Add-in 必须选择一种行为并移除例外。

另有三个旧 Tool 把应用专属能力错误地发送到另一个应用已占用的 Action 名称：Excel 批注使用了 Word 的 `addComment`，Excel 查找替换使用了 Word 的 `findReplace`，PowerPoint 插图使用了 Word 的 `insertImage`。迁移对照将它们标记为 `conflict`，而不是伪装成已完成映射；PowerPoint 已有不冲突的原始 Action `insertPptImage`，另两个能力需要在后续统一 Add-in 工作中解决原始命名冲突。

## Schema 约定

每个 Action 必须包含：

- `action`：原始、大小写敏感的 WPS Action 名称；
- `application`：`common`、`excel`、`word` 或 `powerpoint`；
- `description`：能力说明；
- `parameters`：JSON Schema 风格的对象参数契约；
- `result`：成功时 `data` payload 的 JSON Schema 风格契约；
- `prerequisites`：调用前必须满足的 WPS 状态；
- `risk`：`read`、`write` 或 `destructive`。

风险规则遵循 ADR 0011：只读 Action 无需确认；普通写入可由当前用户请求授权；删除、清空、替换、关闭未保存内容或可能覆盖输出路径的 Action 属于 `destructive`，调用前必须显式确认。

Excel、Word、PowerPoint 和通用代表样例位于 `tests/fixtures/representative-actions.json`，并同时嵌入相应 manifest 项。校验器会用声明的参数和结果 Schema 检查样例。

## 校验

在仓库根目录运行：

```bash
python3 scripts/validate_action_manifest.py
python3 -m unittest tests.test_action_manifest -v
```

校验会拒绝非法或缺失字段、重复 Action、未进入 manifest 的 Bridge dispatch、未映射旧 WPS Tool、指向未知 Action 的迁移映射，以及没有明确说明的 Bridge 缺口或重复 dispatch。
