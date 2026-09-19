# PPT 配色返工根因

2026-09-19。此次仅定因，未修改生产代码或 Windows 安装版。

## 结论

已确认存在 Agent 可见参数契约的说明缺口。颜色参数接受整数，但查询输出未说明编码方式；Agent 使用常见的 0xRRGGBB 整数，执行层则按既有 OLE 颜色规则解释为 0xBBGGRR，导致红蓝通道互换。不是已观察到的 WPS 写色失败，也不是回读返回虚假的成功。

## 可复核证据链

历史证据源：`build/archive/runs/excel-ppt-agent-e2e-r01/analysis/evidence/sessions.json`，数组索引 19 为 PE08 会话，事件索引从 0 开始。

1. 事件 9 是 Agent 当时实际取得的 schema.py 输出。`formatText.parameters.properties.format.properties.color` 只有 `type: integer, minimum: 0, maximum: 16777215`，无编码说明。此结论来自历史工具输出，不是拿当前文件代替当时环境。
2. 事件 28 写入 `req02-format.json`，标题颜色参数为 2046052，即 0x1F3864。
3. 会话中实际解包 PPTX 的工具输出记录标题 `srgbClr` 为 `64381F`，与该整数按 OLE 颜色编码解释的结果一致。
4. 事件 41 写入 `req03-fixcolor.json`，相同标题改传 6567967，即 0x64381F，编码的目标颜色是 #1F3864。该请求另改字号，因此不称为严格单变量实机实验。
5. PE04、PE10 同样记录了首轮蓝色变棕色及后续颜色参数返工；最完整的逐项证据链以上述 PE08 为主。

映射：

| 意图 | 传入整数 | 既有接口实际解释 |
| --- | ---: | --- |
| 深蓝 #1F3864，误用普通 RGB 整数 | 2046052 = 0x1F3864 | #64381F，偏棕 |
| 深蓝 #1F3864，使用正确接口编码 | 6567967 = 0x64381F | #1F3864 |

规则为 `R + G*256 + B*65536`。灰色等 R=B 的颜色不会表现出互换，不能用这些样例证明编码说明正确。

## 缺口位置

`src/main/python/wps_skills/ppt/contracts.py` 中 formatText 的 Action constraints 写有完整公式，formatShape 的 constraints 提到 OLE RGB；setSlideSettings 的颜色参数也缺少解释。

`src/main/python/wps_skills/cli/schema_references.py` 生成 Agent 查询视图时，仅保留 parameters、result、examples，不输出 constraints。因此源码中的解释没有传到 Agent。当前查询结果的以下四个输入字段均没有编码 description：

- formatText.format.color
- formatShape.format.fillColor
- formatShape.format.lineColor
- setSlideSettings.settings.backgroundColor

执行代码直接把整数赋给 Color.RGB / ForeColor.RGB，再回读整数。验证比较的是回读与请求值是否一致，不知道用户心中想要的自然语言颜色。因此可以出现“Action 正确成功，但颜色不符合 Agent 的设计意图”。

## 责任与修复建议

直接触发是 Agent 把颜色编码填错，但不能仅归责为 Agent 不遵守说明：按指定查询流程取得的信息确实不足。应归为接口参数语义没有完整暴露给 Agent。

最小修复是将公式与一个非对称颜色示例放入颜色参数 schema 的 description，通过现有生成流程同步所有四类字段。无需扩大 schema 查询内容，也无需将 COM 实现细节塞进 SKILL.md。参数如何填写属于必要接口信息。

另一种改法是将公共接口改为 #RRGGBB 字符串，由执行层转换；这涉及契约、执行与回读兼容性，不是本次定因必须进行的修改。

复核入口：`python3 build/archive/runs/ppt-color-diagnosis/check.py`。它验证历史查询说明缺失、实际请求参数、历史 XML 颜色证据和当前四字段说明缺失；结果见同目录 evidence.json。本次没有新增 Windows 实机任务，使用已有真实执行记录与当前源码形成证据链。

## 修复完成（2026-09-19）

已将编码公式及非对称颜色示例放入共用 COLOR_INPUT 参数定义，覆盖文字、填充、边框与背景四类颜色。生成查询文件通过 `$defs/COLOR_INPUT` 共享说明，查询脚本自动带回所需定义。输入类型、数值范围、执行及回读行为不变。

本地源目录及新构建安装包的查询检查通过，应用层 11 项测试通过。用户授权后已同步到 `C:/Users/yim/.workbuddy/skills/wps-ppt`，并直接调用 Windows 安装版 schema.py 验证四字段均能解析到完整公式与示例。共更新契约、三个 Action schema，并新增一个公共定义文件。旧文件及安装哈希记录保存在 `C:/Users/yim/ppt-color-fix-20260919`。

本轮是参数契约修复与查询验证，没有重新跑 Agent 端到端测试，不能据此宣称 Agent 再无配色误用。
