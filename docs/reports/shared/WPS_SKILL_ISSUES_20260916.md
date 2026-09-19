# WPS Skill 实测问题与改进方向

**2026-09-17 迁移接续说明：** 下文为迁移前历史计划、讨论和实测记录，不代表当前 Word Task 架构通过验收。当前规则见 [有效 ADR](../../adr/README.md) 与 [ADR 0029](../../adr/0029-make-word-task-the-execution-boundary.md)，实际实施见 [Word 迁移清单](../word/WPS_WORD_TASK_MIGRATION_PLAN.md)。本轮先完成执行侧与现有测试整理，不运行 WorkBuddy Agent 样本，不将历史通过数计入新架构验证。

记录日期：2026-09-16。状态：待处理，未据此修改实现或正式 Skill。

依据：2026-09-15 Windows WorkBuddy 两次用户任务的 SSH 日志核对，以及当前 Word Skill 指引复查。详细原始证据与时间线见 [实测分析](../../../build/evidence/workbuddy-20260915-two-tasks/README.md)。按用户要求先记录在仓库根目录，未创建 GitHub Issue。

## 1. 耗时应如何理解

| 任务 | 用户等待 | Task 工具调用耗时 |
| --- | --- | --- |
| 创建就业现状文档 | 到保存返回 90.696 秒；到最终回复 111.327 秒 | 一次调用 8.247 秒 |
| 更换标题字体、整体排版 | 重新显式指定 wps-word 后，到保存返回 633.860 秒；到最终回复 683.873 秒 | 七次调用合计 42.604 秒，含一次预检拒绝 |

第二次任务主要阶段共 40 次工具调用，调用时长相加约 49.661 秒。相加可能包含并行重叠；这不是严格互斥的墙钟分解。

**已确认：主要等待发生在工具调用之间，WPS Task 执行只占小部分。**

**尚不能确认：剩余时间全部是模型内部推理。** 工具间隔还包含模型请求排队、网络与服务延迟、输入处理、推理、生成文字与 JSON、宿主调度等。现有会话日志缺少足够细的服务端计时，无法拆分各项，也不能给出纯推理耗时百分比。

降低延迟有两个不同方向：减少完成任务所需的模型往返与判断量；降低单次模型响应延迟。本轮已有证据更直接支持先改善前者。不能仅据这次记录断言应换模型，或只优化 Task 执行速度。

## 2. 编排绕路需要计入 Skill 设计责任

上轮报告的“Agent 编排问题”描述的是错误发生的位置，不足以解释错误为何容易发生。当前 SKILL.md、references 与契约呈现方式，把不少本可由产品明确提供的决策留给模型临场推导。它们是需要优先改进的因素；贡献比例仍需修改后的对照实测确认。

### S1：入口偏重 API 操作流程，缺少用户任务到执行路径的映射

现状：SKILL.md 要求先查 Index，再解析完整契约，读若干 references，随后自行编排 JSON。它清楚描述了执行边界，却没有给出“新建并保存”“已有文档改字”“已有文档排版”各自的推荐路径。

实际表现：Agent 花较长时间重新发现能力、选择动作和拼装计划；排版第一次 Task 调用前已等待约 3 分 33 秒。

改进方向：入口先提供少量高频任务的选择依据和通向完整例子的链接；让 Agent 尽快确定需要哪些 Actions、哪些步骤依赖读取结果。保留正式契约作为参数权威，避免另造一套不一致的 schema。

### S2：没有明确区分 replaceContent 的两种分支

现状：content.md 只说使用 range 或 query target，没有突出它们对应不同的 replacement 形状；SKILL.md 又优先强调有匹配数量约束的查找/替换。

正式能力：

- query + text/runs：根据文字匹配替换文字及字符格式。
- range + blocks：按确切范围替换结构化段落，可包含段落格式，且受段落边界等契约约束。

实际表现：Agent 从 query 分支推导出“replaceContent 不支持段落格式”，进而先清空全文，再 writeContent 重写。不是正式能力缺失，而是能力分支没有被便宜、准确地发现。

改进方向：在修改方式指引中明确分支选择，并给出来自 inspect/find 结果的范围引用示例。不能让模型靠阅读大段嵌套 oneOf 来发现常见排版路径。结构替换仍可能改变已有结构和格式，不能宣称它是通用、无损的格式补丁。

### S3：完整 Task 示例与高频保存需求脱节

现状：session.md 唯一完整 Task 示例是 create → write → inspect，并明确故意保持未保存；save/saveAs 需要模型自行拼接。SKILL.md 的批量解析示例同样未含保存。

实际表现：排版请求的 save 步骤漏掉 `params: {}`。session.md 已写明 params 必填，所以不是“从未告知”，而是关键结构没有出现在最容易复用的完整例子里。

改进方向：按需提供完整的新建保存、已有文件编辑保存示例，覆盖空 params、前序引用、必要检查和持久化回执；同时保留“用户未要求保存时可以不保存”的边界。避免把一种任务例子强制套用于所有任务。

### S4：新版 Task 指引与旧式逐 Action 指引混杂

现状：SKILL.md/session.md 已切换到一次命令执行预先编排的计划；verification.md 仍出现“每一步成功并评估后才进入下一步”“Session 仍可用时调整请求”、消费 session.closing/closed、访问 `client.ready` 等低层集成写法。

风险：Agent 需要自行判断哪些步骤由 Task Executor 完成，哪些要自己介入；容易多拆任务，或误以为拿到 Task Response 后仍可沿用原 Session。现有记录不足以证明这次每一次绕路都是该冲突造成，但文档不一致本身可确认。

改进方向：统一普通 Agent 文档任务的 Task Request/Response 工作流。需要语义判断时允许先做读取 Task，再根据结果编排编辑 Task；下一次 Session 必须重新读取可用范围，不能跨 Session 复用 Content Range。只有明确的低层集成才阅读原始 Session 说明。

### S5：验证指引不足以帮助解释当前返回值

现状：指引要求检查 runs 格式，但未解释当前整段聚合的观察限制，也没有区分“属性为 null”“观察范围内格式混合”“写操作精确范围校验失败”。

实际表现：Agent 看到单个 run 的字体 null，认为字体未写入，又把猜测升级为“中西文字体互相覆盖”的限制，写入项目 memory，额外尝试多次修改。

改进方向：先修正底层格式观察粒度；同时提供简短、准确的结果解释。不能只往 SKILL.md 添加“统一字体才有效”等未经证明的 workaround。

另一处缺口：heading 的语义级别不等于视觉样式。首次任务写了 heading，但标题视觉格式仍接近正文；需要明确新建文档的格式选择方式，而不是让模型假定语义标题会自动变大、加粗。

### S6：常见任务读取和契约消费成本高

现状：普通新建保存可能涉及 session/content/persistence/verification 四个参考文件，另加 Index 和多个完整契约；保存说明存在重复。模型还需要从嵌套 schema 中找参数分支和结果引用路径。

实际表现：反复 resolve、Grep、分段 Read；尝试 head 截取输出又因 Windows Bash 缺少命令而失败。

改进方向：SKILL.md 保留最小共同流程和必要边界，按任务类型按需读参考资料；减少重复。评估由同一契约源生成便于 Agent 使用的参数视图或示例，避免手工维护第二份权威定义。不能简单删掉约束或截断契约来换取短输出。

## 3. 需要实现层处理的问题

| 问题 | 已有证据 | 后续工作 |
| --- | --- | --- |
| 等价行距被严格枚举校验拒绝 | multiple=1.5 后读回 oneAndHalf，rewrite 为 unknown | 独立复现，明确规范化和验证规则 |
| inspectDocument 的 run 粒度不准确 | 每段只返回一个聚合 run，无法定位段内格式差异；标题 bold=null，而最终 XML 实际加粗 | 改善按格式片段读取和段落标记处理 |
| 格式修改 changedCount=0 | 字体已改变，文字相同；指纹没有纳入正文格式 | 明确 changedCount 计算，不能以文字未变等同未修改 |

Content Revision 当前定义覆盖内容与结构；是否纳入全部格式需要独立审视领域约定，不能为修 changedCount 顺便改变定义。

Task Executor 在这次记录中正确执行了整份预检、检查失败停止、unknown 停止、保存不隐式执行和清理回执。无需恢复旧的跨命令 Session/后台保活方案。

## 4. 宿主问题与结果报告问题

- WorkBuddy 在未再次显式指定 wps-word 的排版跟进中选了 tencent-docs-routing，用户中断后重新指定。需要检查路由冲突；目前不能断言只改 wps-word description 就能覆盖宿主规则，也没有证据表明已发生云端编辑。
- WorkBuddy Bash 缺少 dirname/head/mkdir，部分命令结尾 echo 还掩盖了失败退出码。需要独立处理环境或使用已验证可用的执行方式；不能把它设计为 WPS Skill 对某一宿主的依赖。
- Session 资源释放不证明用户文档已关闭。最终报告应区分实际观察与推断。
- “全部格式已回读验证”超出了当时声明的 checks。最终文件的目标格式确实已保存，但正确结果不能倒推验证过程完整。

## 5. 建议处理顺序与验收

1. **优先调整 Skill 信息组织**：任务选择、替换分支、完整保存示例、统一 Task 生命周期用语。目标是减少模型必须自行推导的内容，而非堆叠更多“必须”“禁止”。
2. **同时锁定底层格式问题的最小复现**：使用可丢弃文档，分别检查行距、格式片段和 changedCount；不重放用户文档上的未知修改。
3. **对照实测效果**：同一宿主、同一模型、相近输入，区分仅改指引与修复实现的效果。记录首次 Task 前等待、模型/工具往返数、请求拒绝数、Task 耗时、保存返回和最终回复时间，以及结果正确性。单次成功不能证明稳定改善。

期望行为：不再因误读能力而先清空全文；不漏 save.params；不猜段落数；不把 null 当成写入失败；unknown 后先只读核验；最终只报告有证据的结果。

目前尚未修改正式 Skill、运行时、Windows 安装包或 WorkBuddy 的项目记忆。本文件是问题记录和改进方向，不是已完成的修复清单。

## 6. SKILL.md 逐段复查补充

本节针对当前 `src/main/resources/skills/wps-word/SKILL.md`；行号对应本次复查时版本。

### 入口流程的执行时序容易被误读（第 40–52 行）

第 3 节先说“编排完整 Task Request 并执行”，第 4 节才说读取验证说明、验证和保存。按文档顺序执行的 Agent 可能到命令结束后才补充检查/保存，但此时该 Task 的 Session 已清理。正文没有直接说明第 4 节的检查与保存应在第 3 节执行前编入同一计划。第 52 行又统称 save，缺少与前面新建使用 saveAs 的就地衔接。

应按实际工作顺序组织：确定任务与边界 → 选取动作及必要说明 → 编排包含检查、显式保存的完整计划 → 执行 → 解释 Task Response。保存动作仍按新建/已有文件区分。

### “完整任务”“始终复用 Session”的范围不清（第 15、40–42 行）

用户目标可能需要先读取未知内容，再由模型决定如何改写；一个 Task Request 只能执行已能预先确定的区间。session.md 第 24 行解释了这一点，但入口容易让人以为必须把整个用户目标一次编完，并且跨模型决策一直复用 Session。

入口应明确：一份 Task Request 内保持同一 Session；需要模型判断时允许拆成读取计划与后续编辑计划，下一份请求建立新的 Session 和有效范围。这是现有边界的准确表达，不是恢复跨命令 Session。

### 未保存修改的确认规则缺少可执行的检查位置（第 17 行）

规则要求先发现 modified 状态并确认，再开始修改；但文档同时要求一次预编排执行，未说明如何阻止 open 之后直接进入修改。

可以使用现有 checks 在 open 返回后检查保存状态：尚未授权保存原有修改时，以 saved 为继续修改的条件，不满足就停止供 Agent 判断；或者任务本来就需要内容理解时先执行读取计划。已有明确授权时沿用，不应无条件增加一次读取往返或用户确认。此处是指引缺口，尚无证据表明这次任务实际覆盖了用户未授权的修改。

### 入口应提供决策信息，非执行所需细节可后移

第 8、36、60 行混入运行环境内部名称、查询是否启动 PowerShell、测试专用 debug 参数等信息。部分信息对部署和排错有用，但普通文档编排更需要“改文字还是改段落格式”“本次能否预先编排”“保存使用哪个动作”。应优先组织后者，把条件性细节放到对应参考资料，保留路径、支持环境与关键安全边界。

这些问题不能只用“SKILL.md 太长”概括：入口约 62 行，关键在信息顺序、决策缺失与引用链的一致性。后续应以 Agent 是否能更少往返地选对动作作为验证，而不以删掉多少行作为成功指标。

## 7. 后续实施：按名称读取静态 Action 定义

2026-09-16 按用户确认完成 Word Skill 的这一项调整：SKILL.md 列出 14 个 Action 的名称和用途，Agent 选择动作后从 `references/actions.json` 读取完整定义，再补齐编排；正常指引不再要求 `--index` 或 `--resolve`。CLI 参数仍兼容保留，Excel/PPT 的发现流程本次未改。

该文件以 Action 名称为顶层键，从生产契约生成；`python scripts/build/word.py --refresh-actions` 同步源码参考文件，构建时自动重新生成并纳入哈希清单。相关 references 已同步。其余前述格式实现问题和 Skill 流程问题仍待处理。

验证：16 项构建/CLI 测试通过；Skill 校验通过；文档中的提取示例在独立构建包、其他工作目录下成功读取所选四个完整定义；14 个动作与正式契约一致，44 个打包文件哈希通过。新包在 `build/archive/packages/skills-actions-20260916/wps-word`，未替换 Windows 安装包，未新增真实 WPS 文档操作。

## 8. 后续实施：精简 Agent 阅读材料

2026-09-16 按用户要求继续清理 Word Skill：根据名称和用途直接编排动作顺序，填写调用参数时再按需读取完整定义。SKILL.md 增加新建保存、已有文件编辑、只读任务的常用动作顺序。

删除原 references 中的 session.md、logging.md、content.md、persistence.md、verification.md。面向调用者的必要内容合并到 task.md：任务 JSON、前序结果引用、检查条件、返回结果和不确定修改的处理；提供包含保存的完整示例。正文与 task.md 不再讲 Session、日志位置、COM、进程、锁或调试配置。references 现仅有 actions.json 和 task.md，完整契约仍按代码生成；运行时行为未改。

同步 README 与 INSTALL 的相关链接。Agent Markdown 总字符数由上一轮构建的 13,722 减至 6,188，约减少 55%。3 项 Word 构建测试、Skill 校验通过；新示例通过实际任务结构与 Action 参数预检，引用及检查条件核对通过；包内链接和 40 个文件哈希通过。新包为 `build/archive/packages/skills-word-guide-20260916/wps-word`，未更新 Windows 安装，未执行实际文档修改。其余排版实现问题继续保留待处理。

## 9. 后续实施：通用调用说明并入 SKILL.md

按用户反馈，task.md 的请求格式、引用和结果处理是所有执行都需要的材料，单独拆分增加了阅读跳转。已合并至 SKILL.md，并合并重复说明、删除 task.md；references 仅保留按操作名称读取的 actions.json。README、INSTALL 链接同步更新。

合并后的示例通过真实请求和 Action 参数预检，结果引用与检查条件验证通过，按名称提取定义的命令在其他工作目录可用；Skill 校验、14 个名称匹配、文档链接及 39 个打包文件哈希通过。新包为 `build/archive/packages/skills-word-unified-20260916/wps-word`，未更新 Windows 安装。

## 10. 后续实施：Word Task 生命周期配置、直接提交与统一注册表

2026-09-16 按用户确认落实三个接口调整，修改范围为 Word 新请求入口及其共享 Task 执行支持：

- Word Task v2 使用 `document`、内容 `steps`、`completion`。执行器按固定顺序获取／复用精确文档，执行内容步骤和检查，再保存、导出 PDF。Agent 不再在 steps 中安排 create/open/save/saveAs/exportPdf；这些调用仍有独立执行结果，失败或 unknown 不会继续收尾。
- `completion.mode` 支持 none/save/saveAs，PDF 单独配置；保存已有文档时，若发现原有未保存修改且未明确授权一起保存，在内容步骤前停止。读取请求不保存，也不要求该授权。保存成功但 PDF 失败时保留 Word 已保存的事实。
- 正常使用 `--task-stdin`：一次命令直接提交 UTF-8 JSON，无需调用者创建 task.json。修正 Windows 非 UTF-8 默认编码可能造成的标准输入解码问题。内部回执继续保留，以支持状态查询、冲突检测及避免重放；这与要求调用者先创建任务文件不同。
- `word/registry.py` 集中关联全部 14 个 Action 的 contract 和 handler，并派生正式契约集合与兼容 handler 视图。获取文档也通过注册项分派。Word SKILL.md 和 actions.json 仅暴露 9 个可编排的内容操作，保存与导出通过配置表达；底层 14 项能力不删减。
- Task Response 分开返回 document、steps、completion.save、completion.pdf，未请求的收尾项为 null，失败后尚未执行的阶段为 not_executed。v1 请求、--task-file 及其他应用旧接口继续兼容。

SKILL.md 已同步目标／交付配置、内容操作选择、直接提交示例和响应处理；README、INSTALL、脚本说明已同步。CONTEXT 与 ADR 0024 记录新边界，旧 ADR 0004/0013/0023 标记为被替代；这些领域文档按仓库现有规则属于本地忽略文件，未更改忽略配置。

验证：完整本地单元／集成套件 259 项通过；后续增加 PDF-only、保存成功但 PDF 失败等覆盖，并在最终实现上运行相关回归 36 项通过（包含 13 项 Word Task v2 测试）。覆盖动态范围引用、固定执行顺序、原有修改确认、失败跳过保存、保存不确定后不重放、生命周期中的进程丢失、真实 CLI 标准输入及中文／emoji。独立打包后的 v2 入口可在没有 WPS 的环境下正确预检并拒绝无效保存路径，未启动文档操作。SKILL 示例通过真实结构及参数预检，源文件与构建包的 Skill 校验、9 个内容名称／schema 一致性、链接和 41 个文件哈希检查通过。

新包：`build/archive/packages/skills-word-task-v2-20260916/wps-word`。尚未替换 Windows 已安装 Skill，未做本次改动的真实 WPS 文档验收。此前记录的格式读回、changedCount 等实现问题不属于本次修复，仍待处理。

## 11. 用户澄清后的修正：三部分统一为 Action 请求

用户指出固定头尾执行位置不等于隐藏获取和交付 Actions。已移除第 10 节中的 document.kind、completion.mode 配置别名：document 为一个 address + params 请求，steps 为带 id 的内容 Action 列表，completion 为 0–2 个 address + params 请求；各部分均支持 checks，并复用原始 Action schema。执行器限制 document 只能选择 createDocument/openDocument，completion 最多一个 save/saveAs 和一个 exportPdf，并固定获取在前、保存与导出在后（先保存再导出，不依赖 completion 输入顺序）。参数不再被配置映射补齐。

原有未保存修改授权 includeExistingChanges 移至 Task Request 顶层，避免混入 Action params。文档获取、保存、PDF 的固定响应 id 及分阶段响应保持不变。新版 v2 尚未部署，直接修正其接口；v1 兼容保持。

SKILL.md 恢复完整 14 个 Action，分 document/steps/completion 三组说明名称、用途；actions.json 从统一注册表导出全部 14 份契约。同步示例、README、安装与脚本说明，ADR 0025 替代 0024 记录此次澄清。新增获取 Action 检查、保存检查阻止 PDF、非法分组／重复收尾／缺失原始参数拒绝等覆盖。47 项相关回归通过，Skill 校验、文档示例真实预检、14 项名称及 schema 一致性、41 个包内文件哈希通过。

最新包：build/archive/packages/skills-word-action-groups-20260916/wps-word。未更新 Windows 安装，尚未执行真实 WPS 验收。

## 12. 三部分 Action 结构的首次 WorkBuddy 实测

详见 `build/evidence/workbuddy-20260916-action-groups/README.md`。Windows 安装包 41 文件哈希一致；6 次工具调用、2 次 Task，三部分请求和 --task-stdin 使用正确。第一轮创建/写入/检查成功，末尾 saveAs 因同名文件 OUTPUT_ALREADY_EXISTS 失败；Agent 换名后创建并写入第二份相同内容，最终成功保存 `中国毕业生就业现状2026.docx`。两轮清理均成功，无 unknown 或 Session 丢失。

总耗时 98.752 秒，两次 Task 命令合计 15.995 秒；所有工具区间去重后 17.014 秒，剩余 81.738 秒不能仅解释为纯推理。第一轮失败返回到第二轮成功用时 29.507 秒。优先问题是将已确定的输出路径冲突提前到写入前检查；同时明确资源释放不等于未保存文档丢失。当前接口不能跨请求重新取得无文件路径的新文档，不能假设可以直接接回原对象保存。

用户要“一段”，实际生成五节、11 段、995 字符；标题结构正确但只设居中，读回 10.5 磅且不加粗，小于 12 磅正文。完整文字与请求逐字匹配；原 checks 仅包含关键字和未截断，没有全部格式断言。schema/两轮工具回显仍有较大上下文开销，具体数据见报告。此次只分析日志，未修改实现、重放 WPS 操作或改变 Windows 安装。

## 13. 后续实施：收尾 Action 内部解决输出名称冲突

按用户最新确定的边界，本次不实现 Task 重试、失败挂起或跨请求接回未保存文档。Task 最终失败后仍关闭 Session；已完成修改不回滚，资源释放不表示文档丢失。第 12 节建议的独立输出预检本次未新增，改由收尾 Action 的原子检查与有界换名处理冲突，避免预检后仍可能出现的竞争。

- Word saveAs、exportPdf 增加 overwritePolicy: renameIfExists。先尝试原路径，再在同目录、同扩展名下尝试“名称 (1)”“名称 (2)”等，最多 20 个候选；始终保留同一文档与原 Action 截止时间，不重复获取或内容操作，不需要 Agent 介入中间尝试。
- 只重试 outcome=failed、绑定 unchanged 的 OUTPUT_ALREADY_EXISTS 或 OUTPUT_IN_USE。unknown、超时、权限错误、文档身份变化、Lease 冲突及 Quarantine 不自动重试。候选耗尽返回 OUTPUT_NAME_EXHAUSTED，其他错误保留原结果并停止。
- 返回实际 artifact.path 及 outputResolution.requestedPath/attempts/renamed；契约检查候选路径、次数和未覆盖事实。普通 save、严格 failIfExists、PDF replaceExisting 不增加自动重试。
- 保存继续使用原子 CreateNew 占位，并将占位瞬间的同名竞争归类为明确冲突。PDF 发布改为无覆盖文件移动，防止预检后新出现的同名文件被 Move-Item -Force 覆盖。
- SKILL.md 推荐普通另存／导出使用 renameIfExists，必须原名时使用 failIfExists；明确按实际返回路径交付。更新完整 actions.json、README 及 ADR 0026（替代 0021，保留原有持久化保证）。

验证：本地完整套件 270 项通过；新增 7 项重试／契约测试。Windows 隔离包真实 WPS 验收预先占用 DOCX/PDF 的原名和 (1)，两个 Action 各尝试 3 次并使用 (2)，四个原有哨兵文件均未变化，DOCX 正文只写入一次，Session 日志只记录一次 createDocument/writeContent/inspectDocument/saveAs/exportPdf。额外 PowerShell 模拟 PDF 导出期间产生同名目标，验证不覆盖、返回明确冲突、临时文件清理。源 Skill 与最终包校验、示例真实预检、14 份 schema 一致性、42 文件哈希检查通过。

最终包：build/archive/packages/skills-word-output-retry-final-20260916/wps-word.zip。运行时代码与 Windows 实测包逐文件一致，仅修正了最终 SKILL.md 的一处策略说明。详细证据：build/evidence/word-output-retry-20260916/README.md。Windows 只使用隔离临时目录，未更新 WorkBuddy 安装。

## 14. 第一阶段问题修复：等价行距与 changedCount

第一阶段在真实 WPS 复现 multiple=1.5 实际规范化为 oneAndHalf，而验证直接比较 kind 导致 unknown。现改为按实际倍数比较 single/oneAndHalf/double 与 multiple 的等价表示；exact/atLeast 继续区分，真实效果不一致仍拒绝。未改输入格式、公开读回类型或不确定结果后的停止策略。

用户随后明确选择删除 replaceContent、setHeaderFooter、setPageLayout 三处对外 changedCount。该字段原用于变化目标数量摘要和 revision 一致性检查，但不能表达纯格式变化。现删除对应 schema 和原生返回，保留 matchedCount、selectedSectionCount、fresh ranges、stories/sections、前后 revision 及内部效果回读；删除/不同文字替换仍要求 revision 变化。页眉页脚与页面布局内部使用变化布尔量继续验证，不将 Content Revision 扩展为完整格式版本。

旧调用者若在 checks 或结果引用中使用 changedCount，需要更新；新 schema 与运行时须作为完整包使用。源生成 actions.json 已同步；SKILL.md 无该字段说明，未添加一次性兼容规则。

验证：旧包行距 24 场景中 6 个等价互换失败，新包 24/24 通过；本地全套 274 项成功（1 个 Windows 专用项跳过，底层测试已在 Windows 独立通过）；真实 WPS 六种行距写法及三类结果联合场景共 7 例通过，保存 DOCX 与截图核验正确。新包 42 文件清单及 Skill 校验通过，测试文档/桥接进程残留零，原有文档保留。

新包：build/archive/packages/skills-word-line-spacing-fix-20260916/wps-word.zip；SHA-256 `b833b4e9e692d6a0e37abddaed12465be43d112629b98d94b8d624fa92e0134e`。详见 [修复与回归报告](../../../build/evidence/word-line-spacing-fix-20260916/README.md)。未替换 WorkBuddy 安装，也未启动第二阶段测试。

## 15. 第二阶段：真实桌面自动提交链路及首个 Skill 样本

2026-09-16 后续，修复包已传 Windows 桌面，用户自行安装并授权完整自动提交。ZIP 全部 43 文件与安装内容一致（含 42 文件的哈希清单自身）。Codex 经真实 WorkBuddy 桌面新建会话、填写自然语言、点击发送、采集日志并独立验证产物，全程无需用户粘贴或构造请求。没有改生产实现或 WorkBuddy 配置。

- **ST2-01 路由**：不点名 Skill 的 E01 被宿主 `tencent-docs-routing` 导向内置编辑技能；文件文字和单段目标通过，但不能计为 wps-word 样本。明确点名后 E01B 成功加载 wps-word。
- **ST2-02 编码**：模型虽然生成了正确 UTF-8 JSON，却遗漏 Skill 示例中的 `$OutputEncoding` 设置。PowerShell ASCII 管道把正文和检查 expected 同时变成问号，contains 因而仍通过。执行器持久请求已损坏；默认管道连续两次复现失败，显式 UTF-8 对照通过，均未调用 WPS。这是请求传输/Agent 使用问题，不是 WPS 改坏正确输入。
- **ST2-03 结构与断言**：写入一个 paragraph 后，inspect、WPS COM 与最终 OOXML 都是两个段落，第二段为空。前两次 Task 因 paragraphCount=1 检查停止；模型第三次删掉数量断言后保存。不能把删断言称为严格用户目标得到修复。
- **ST2-04 重建**：模型两次换 ID 再次 createDocument，最终产生三份文档，两份未保存。它识别到没有文件路径无法重绑，却仍重建，违反只创建一份的目标，也偏离 Skill 保留已发生效果、不重新创建整篇的指导。没有 unknown，三次 Session cleanup 均 succeeded；资源释放成功不等于没有多余文档。
- **ST2-05 宿主工具**：PowerShell 多次只返回退出码，模型依靠落文件再 Read 取得结果；Bash 有 dirname/cd 环境错误。当前属于已观察但未独立定位的宿主问题，未为此改包或放宽权限。
- **ST2-06 交付**：模型报告“无其他段落”，实际多出空段落；独立 XML 验证只拼接非空文本 run，无法证明段落数量。它还在自己的测试 workspace memory 写入“不应断言 paragraphCount=1”的概括，后续样本需注意上下文影响；本轮未修改这份记忆。

E01 桌面会话 208.25 秒、31 次工具；E01B 343.28 秒、28 次工具、3 次 Task。E01B 最终 DOCX 文字完全匹配、原生 Saved=true，但严格用户目标未通过。本轮只证明完整自动提交链路在当前桌面状态可执行，不宣称批量稳定性或所有设计通过。两个已知测试未保存文档仍留在 WPS，最终产物也保持打开；13 个临时交互计划任务已清理。

完整原始记录、三份请求/回执、编码对照、原生观察与两份产物见 [完整链路报告](../../../build/evidence/workbuddy-e01-20260916-1852/README.md)。后续需分别处理输入编码、段落结构、失败恢复与宿主路由；本轮尚未实施这些修复。

## 16. UTF-8 入口与尾端空段修复（2026-09-17）

本轮用户确认由固定 PowerShell 提交脚本接收 Agent 的 JSON、设置 UTF-8 并启动 Python；无需 Agent 编写编码配置或生成传输文件。已增加 `scripts/submit.ps1` 并更新 Skill 和安装说明。Windows 回归覆盖 ASCII/GB2312/UTF-8、中文/emoji、长 JSON、空白行、stderr 进度及退出码。

`writeContent` 在文档末尾省去自己生成的最后一个 CR，复用 WPS 必须保留的最终段落标记；同步收缩返回范围和格式计划，保留文中分隔及有意写入的空段。真实 WPS 基线可重现额外空段，修复后的新建、追加、文中插入、格式、范围复用和持久化回归通过。补查发现非空末段后插表格时 WPS 自动增加分隔符，表格起点相应后移一位；已针对 `documentEnd` 同步修正位置校验。不修改通用检查比较器，不将段落数量检查一概加一，也不改写第 15 节历史证据。

测试范围、清理结果与产物见 [回归报告](../../../build/evidence/word-submit-tail-20260917/README.md)。WorkBuddy Agent 是否稳定采用新入口仍需独立端到端验证；保存失败交付语义不在本次修改范围内。

## 17. 新包安装后的桌面复测（2026-09-17）

44 个安装文件与新包完全匹配。真实 WorkBuddy Agent 的单次 Task 生成 1 段正确正文并保存，原生 COM 和 DOCX 双重核验通过，没有重复创建；尾段修复在宿主编排场景得到验证。

ST2-05 的 PowerShell stdout 缺失仍存在。Agent 用输出探针确认宿主只返回退出码后，绕过未执行的 `submit.ps1`，写入 JSON 文件并用 Bash 执行兼容的 `--task-file`。因此无文件提交验收未通过，不能称编码入口已在 Agent 链路验证成功。Agent 还将该替代方式写入自己的工作区 memory；这不是本仓库的新决定。详见 [安装后复测报告](../../../build/evidence/workbuddy-utf8-tail-20260917/README.md)。本轮不更改产品代码。


## 18. 改用可消费的 task file（2026-09-17）

用户基于第 17 节实际宿主表现，决定使用 task file，并选择持久化接收后自动删除临时输入。推荐使用 UTF-8 JSON 文件和 `--task-file PATH --consume-task-file`；原始请求和初始回执先原子写入并刷新到磁盘，再删除临时输入，随后继续执行。执行失败和中断也保留按 taskId 保存的证据。未接收、ID 冲突保留输入；清理失败单独报告，不改变 Task Outcome，也不触发重建。普通 `--task-file` 保持非消费语义。

移除未被宿主采用的 `submit.ps1`，同步 Skill、安装说明及 ADR 0027。此前尾段 CR 和表格末尾插入修复保留。本地与 Windows 各 47 项回归通过；新包已构建，尚未进行安装后的 WorkBuddy 自主编排验收。范围及证据见 [本次报告](../../../build/evidence/word-task-file-20260917/README.md)。


## 19. task file 安装后实际链路通过（2026-09-17）

用户安装后授权继续测试。WorkBuddy 在未获得启动命令提示的情况下自主采用 `--task-file … --consume-task-file`，临时 JSON 删除，原始请求和执行回执均保留，`--task-status` 查询与归档一致。一次创建、一次保存；完整文字、段落数为 1、未截断三项检查通过，原生 WPS COM 和 OOXML 双重验收通过。64.65 秒、11 次工具调用；44 个安装文件测试前后均匹配，没有产品修改。四个临时桌面计划任务清理完毕。详见 [验收报告](../../../build/evidence/workbuddy-task-file-20260917/README.md)。

ST2-05 的 Bash 缺少 ls/dirname 噪音仍可观察，但未阻碍本次完成；PowerShell stdout 本轮没有调用验证。单样本通过，不扩展为全部宿主稳定性结论。


## 20. 第二批 E01/E02/E04 各三次完成（2026-09-17）

9 个固定样本已跑完（复用同版本 E01/r1，新增 8 例）。9 例正文和保存效果符合目标，7 个完整通过，2 个编排/交付问题。E02 三例均两处精确修改、格式和备份保留、原文件保存；E04 三例均单次 saveAs 内部 3 次候选后保存 `(2)`，两份已有文件哈希不变，未重建。

| 编号 | 观察事实 | 当前结论 |
| --- | --- | --- |
| B2-01 / E01 r2 | Agent 最终称 Session 释放意味着文档关闭；最后回复后的 COM 与截图确认文档仍打开且已保存 | 用户后续确认不作为问题处理，已移出待处理项；原始观察保留，未改实现 |
| B2-02 / E02 r2 | 复用 r1 的 taskId，被 TASK_ID_CONFLICT 在接收前拒绝；换新 ID 后完成 | 身份保护正常，编排有额外拒绝和等待；被拒绝输入保留符合生命周期 |
| B2-03 / E04 r1 | saveAs 成功后，checks 的路径误加 response 层，TASK_REFERENCE_UNAVAILABLE 导致 Task failed；没有重建，最终披露检查失败并交付正确文件 | 编排错误；task_save 可引用，正确路径从 data 开始；memory 却错误归因为保留 ID 不受支持，尚未修正 |
| B2-04 / E04 r2 | PowerShell 六次调用仍只返回退出码，Agent 通过导出文件/读取回执等绕行，15 次工具、199.25 秒完成 | 既有宿主 stdout 问题持续，不影响本次文档效果但增加成本 |

完整原始请求、失败回执、工具调用、独立原生/XML/哈希核验与耗时见 [第二批报告](../../../build/evidence/word-batch2-20260917/README.md)。4 例超过 120 秒观察线，不以此单独判功能失败；未测首次肉眼可见进展和性能因果收益，不估计长期成功率。

安装前后 44 文件一致；本轮 11 个新 Task 的 Session 均成功清理，56 个临时计划任务删除，8 份新测试文档核实保存后关闭、产物保留，原有文档不动。中途自动审批以远程授权不足拦截一次，用户明确“给权限”后恢复，未重复发送样本。未改生产代码或 Agent memory，未启动第三批。本节为本地测试观察台账，尚未发布 GitHub Issue。


## 21. Word Task v3：身份、验证与生命周期收进代码（2026-09-17）

用户确认 B2-01 不作为问题处理，并要求普通 Word Action 依靠已有内部回读，不再由 Agent 编排 checks；Task 对外统一，Session 资源生命周期归代码。已实现 v3：请求不接受手填 taskId/checks，代码按独立输入路径分配任务身份，自动持久化接收与消费输入，删除后仍可用原路径查询，重复提交不重放，修改已接收路径拒绝。参数结果引用保留，内置 Action 验证与已有未保存修改保护保持有效。Word Skill 入口拒绝旧版计划；共享低层集成仍保留旧接口。回执只暴露 Task 身份与清理摘要，内部 ready/session/stderr 归诊断文件。

最终本地 297 项通过（1 跳过），Windows Word 专项 57 项通过；最终隔离包真实创建/修改/重名保存和原生 COM、DOCX、截图验证通过。44 文件一致，新包已放 Windows 桌面，未替换安装版本。原生测试夹具两处参数/目标误用已修正，早期测试文档仅补做保存后关闭，未重放修改；首次 v3 暴露 ready 的产品遗漏也已修复并回归。详见 [v3 报告](../../../build/evidence/word-task-v3-20260917/README.md)。安装后 WorkBuddy Agent 复测尚待用户安装，不将本轮直接 CLI 验收算作第二批 Agent 重跑。
