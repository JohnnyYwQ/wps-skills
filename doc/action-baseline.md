# WPS Action 基线

Issue #3 冻结了迁移开始时的能力事实。机器可读文件是唯一依据；本文只解释边界和审计结果，不重复维护 Action 列表。

## 事实来源

- `skills/wps-office/references/action-manifest.json`：公开 WPS Action 契约。每项记录应用归属、参数 Schema、结果 Schema、前置条件和 `read` / `write` / `destructive` 风险等级。
- `doc/migration/legacy-tool-action-map.json`：冻结的一次性迁移对照。它记录全部旧 WPS Tool 到原始 WPS Action 名称的映射和明确退役项，作为删除旧运行时后的审计证据。
- `scripts/validate_action_manifest.py`：Python 3.9+ 标准库校验入口。它只读取最终 manifest、打包的 JavaScript Add-in 与冻结迁移对照。

公开接口只使用 manifest 中的 `action`，例如 `getCellValue`。旧 WPS Tool 名称只存在于迁移对照中。

## 冻结时的审计结果

| 范围 | 数量 | 说明 |
|---|---:|---|
| WPS Action manifest | 249 | 最终公开的规范 WPS Action |
| 打包 JavaScript Add-in dispatch | 249 | 与 manifest 中的 WPS Action 逐项完全一致 |
| 冻结旧 WPS Tool 对照 | 250 | 包含映射、工作流映射与明确退役项 |

`addArrow` 统一采用起止坐标绘制线箭头；先前的边界框块箭头形式已按 ADR-0014 退役。manifest 只接受 `startX`、`startY`、`endX` 和 `endY`，迁移对照不再保留契约冲突例外。

旧 WPS Tool 中曾有应用专属能力错误发送到另一个应用已占用的 WPS Action 名称。迁移对照现在将 Excel 批注修正为 `addCellComment`，将 PowerPoint 插图修正为 `insertPptImage`；错误复用 Word `findReplace` 的 Excel 查找替换则按 ADR-0014 明确退役。不存在未决的 `conflict` 映射。

## Schema 约定

每个 WPS Action 必须包含：

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
python3 -m unittest discover -s tests -v
```

校验会拒绝非法或缺失字段、重复 Action、未进入 manifest 的打包 Add-in dispatch、manifest 中缺失的打包 Add-in dispatch，以及指向未知或跨应用 WPS Action 的迁移映射。
