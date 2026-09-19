# Word Session → Task 迁移清单

日期：2026-09-17。状态：执行侧与测试整理已实施，本地回归通过；Windows 原生验收待执行。

本清单依据已确认设计实施。第 2 节保留迁移前差距作对照，第 3–6 节记录范围与门槛；实际改动、测试处置和验证结果见第 8 节。用户已授权修改 Word 相关代码、测试和文档；本轮不运行 WorkBuddy Agent 样本。

## 1. 已确认的迁移边界

| 项目 | 本轮结论 |
| --- | --- |
| 应用范围 | 只迁移 Word。Excel、PPT 后续再统一，当前继续使用原有执行链。 |
| 编排与提交 | Agent 理解自然语言、发现 Action 能力，选择 `document`、编排 `steps`、选择 `completion`，将 Task Request 交给 Task Client。 |
| Task 边界 | 一次提交的执行计划。Task Client 是调用入口；Task 是执行和资源管理边界，两者不另建两套生命周期。 |
| 生命周期 | Word 原 Session 管理的文档绑定、Lease、执行状态、进程与资源释放全部由 Task 接管。Word 不再通过内部 Session 执行。 |
| Action 职责 | 执行具体操作并验证实际效果，向 Task 返回结果。Task 根据结果继续或停止，不替 Agent 修改计划。 |
| 请求结构 | 删除 Task Request 的 `version`，仅支持当前结构；不接受调用方 `taskId` 或 `checks`。保留内容步骤 ID 和前序结果引用。 |
| 保存默认值 | 新建和修改已有文档都默认不保存。用户未要求保存或导出时使用 `completion: []`。 |
| 固定执行顺序 | 获取文档 → 已编排的内容操作 → 被选择的交付操作；保存先于 PDF 导出。没有选择的操作不补做。 |
| 异常 | 停止后续操作，保留已完成效果，不回滚、不自动重放。Agent 根据回执决定新的编排；效果不确定先核实。 |
| 保存失败 | 报告用户、保留现场、结束 Task，不触发 Agent 补救，不标记为全部成功。保存策略及操作验证留在保存 Action 内。 |
| 结束 | 立即进入有界资源释放，不等待 Agent 的下一次编排；不额外保存、关闭文档或退出 WPS。无法证明底层操作停止时保留必要的 Quarantine。 |
| 旧接口 | 移除 Word 旧 Task 协议和 Session 工程入口，不为 Word 保留兼容执行层。 |
| 文档 | 更新有效决策和领域定义，当前阅读入口只指向有效 ADR；历史标记已取代，设计演进记录串起调整和原因。 |

**保存意图与执行分开处理。** Skill 指导 Agent 从用户要求中选择 `completion`；执行器不猜自然语言意图，也不新增一套自然语言判断器。执行器必须保证空 `completion` 不触发保存、导出或清理保存。

本轮不新增未保存文档的跨 Task 恢复地址、常驻会话、自动补救服务或新文档能力。请求中既有的 `includeExistingChanges` 与原有未保存修改保护不因默认保存规则变化而取消：它只在明确选择保存已有文档时适用。

## 2. 迁移前代码与目标之间的差距（基线）

当前 Word 主链路为：

```text
word.py → word.skill.main → cli.call → task_client
  → SessionClient → CLI --session → SessionHost
  → Word build_session → ActionSession / Runtime → Adapter / Bridge → WPS
```

迁移后的职责链为：

```text
Agent → Task Request → Task Client → Word Task 执行环境
  → Action / Adapter / Bridge → WPS
  ← Action 结果 ← Task Response ← Agent
```

任务内仍可复用同一个受控桥接进程和精确文档引用。进程、传输、参数及结果验证是实现机制，不构成第二个可独立开始、等待和结束的 Session。

Session 原有的交互式等待、空闲等待超时及 `session.ready/closing/closed` 协议不需要照搬到完整计划执行中；Task 仍须保留操作超时、取消、调用进程丢失和有界清理。Windows 的桌面登录 Session 是操作系统概念，与本轮移除的 Action Session 不同，不能机械删除名称中所有 `session`。

已核实的主要差距：

- `client/task_client.py` 仍调用 `SessionClient`，并以 `session_outcome` 等字段判断清理。
- `word/task_plan.py` 把三部分请求编译成带旧 `version: 1` 和占位 `taskId` 的计划；不能只删外层 `version`。
- `task_request.py`、`task_store.py`、`cli/call.py` 多处以版本号决定结构、身份、输入消费与状态查询。
- `core/action_session.py` 同时容纳通用类型、契约验证、绑定资源聚合、Runtime 和 ActionSession；Word 的 contracts、adapter、handlers、诊断仍依赖它。
- `word/skill.py` 和 Skill 的 `scripts/word.py` 仍导出 `open_session` 及 Session 异常类型。
- `build_skill.py` 当前整包复制 `core/client/host/windows`，即使入口改完，Word 安装包仍可能携带旧执行栈。
- Word 的部分测试、演示及原生验收仍使用旧请求、Session 夹具、Agent checks 或 stdin 提交。

## 3. 代码迁移清单

源码路径以下均相对 `src/main/python/wps_skills/`，资源路径另行标明。新模块名称可在实现时按职责确定，本清单不要求为每行新增一个模块。

| 编号 | 现有位置 | 后续处理 | 完成条件 |
| --- | --- | --- | --- |
| C01 | `word/skill.py`；`src/main/resources/skills/wps-word/scripts/word.py` | 建立 Word 当前 Task 提交、查询、能力发现入口；移除 `open_session` 及 Session Client 类型导出。 | Word 的安装入口和源码入口都不能启动旧 Session 或提交旧计划。 |
| C02 | `cli/call.py`；`client/managed_session.py` 的 Word 路径 | 按应用分开执行路由；Word 拒绝 `--session`、`--start/--call/--status/--close`、旧 stdin Task 和旧消费开关等入口；删除 Word Session 工厂分支。保留 Task 回执查询和无副作用的 Action 发现。 | 不可从共享 CLI 绕回 Word 旧接口；拒绝发生在启动 WPS 和文档操作之前。Excel/PPT 原入口仍有效。 |
| C03 | `word/task_plan.py`；`client/task_request.py` 的 Word 路径 | 直接验证无 `version` 的三部分请求，生成当前内部计划。移除 Word v1/v2/v3 分流、checks 执行和编译到旧协议的转换。 | 开始和结尾只能选择合法 Action；步骤顺序、ID、引用和参数受校验；`version`、调用方 `taskId`、`checks` 均不能被静默忽略。 |
| C04 | `client/task_client.py` | 将 Word 执行与 Excel/PPT 的 Session 执行分开；Word Task 自己取得执行环境、执行并记录 Actions、停止和释放资源。Task Client 不再为 Word 启动 `--session` 子进程。 | 一份请求只有一个 Task 生命周期；没有把 Session 改名、包装或继承为 Task 的过渡执行路径。 |
| C05 | `core/action_session.py` 的通用类型、验证函数及 Runtime 机制 | 抽出仍需复用的契约、命令/结果、验证和资源机制，解除 Word 对 Session 模块的依赖。保留 Excel/PPT 所需 Session 实现，避免复制整套 Core。 | Word 使用的公共机制没有独立 Session 身份或生命周期；Excel/PPT 可继续复用通用机制。 |
| C06 | `core/action_session.py` 中绑定、Lease、清理机制；`word/adapter.py`；`windows/document_coordinator.py` | 将 Word 的逻辑状态、资源聚合和清理所有权归到 Task；Adapter/Coordinator/Bridge 执行实际资源操作。 | 绑定和 Lease 在成功可见前提交；一个 Task 只绑定一个文档；获取中失败、部分获取、正常结束、取消、异常均由同一 Task 清理出口处理。 |
| C07 | `word/windows/session.py`；`word/windows/backend.py`；`windows/bridge_runtime.py`；`windows/owned_process.py` | 用 Task 的 Windows 组装替换 `build_session`，迁移桥接位置和启动错误定义；删除 Word 专属 `windows/session.py`。继续复用进程所有权、超时和桥接机制。 | Word 无 SessionHost/SessionClient/ActionSession 依赖；每个 Task 复用自身桥接资源，Task 之间不共用活动文档状态。 |
| C08 | `word/handlers.py`、`word/contracts.py`、`word/registry.py`；`word/output_names.py`；Word PowerShell 资源 | 调整通用类型导入和执行上下文来源；保留 Action 参数、结果及原生效果验证。保存失败由 Action 明确报告，Task 不另加保存重试循环。 | Action 始终操作 Task 的精确文档；保存冲突候选选择仍在同一个 Action 内，失败不重建文档、不重新编排内容。 |
| C09 | `client/task_store.py`、`client/task_file.py`；`task_client.py` | 按当前 Word 契约记录原请求和回执，移除运行时对请求版本的依赖。将 JSON 读写/解码等从 `managed_session.py` 分离为真正通用帮助函数。 | 保留输入路径身份、持久接收后消费、原输入路径查回执、并发互斥、相同请求不重放、冲突输入保留。旧磁盘证据不删除或重放。 |
| C10 | `task_client.py`、`task_store.py`；`core/trace_journal.py` | Task Response 如实记录各阶段、停止原因、保存结果及清理结果；诊断关联 Task 与 Action，移除 Word 的 Session 身份、握手和结束回执依赖。 | 保存失败可直接识别且不报全成功；未知效果不改写为确定失败；清理异常不抹掉已确认效果。响应/存储结构不再靠旧请求版本分流。 |
| C11 | `src/main/resources/wps_skills/word/windows/word_bridge.ps1`、`word_com.ps1`、`word_actions.ps1`；共享 Windows 资源 | 校正 Word 内部资源归属、注释和诊断命名，检查 Revision 的来源和生命周期。继续保留必要的 COM、文件身份、Lease 和 Quarantine 实现。 | Content Range 仅在生成它的 Task 内有效；跨 Task 重新读取。清理释放引用及自动化资源，不保存或关闭用户文档。 |
| C12 | `word/demo.py`、`word/desktop.py`；`scripts/demo/word.*`、`scripts/validate/word.py`；混合持久化夹具的 Word 分支 | 将仍有用途的演示和验收改为正式 Task；移除 Session 演示、专用旧协议脚本及过时前提。 | 演示不再通过旧入口维持兼容；工程夹具的显式清理不混入产品 Task 的默认行为。 |
| C13 | `cli/build_skill.py`、`cli/build_word_skill.py` | 按 Word 实际依赖打包当前执行链；保留其他应用的打包路径；重新生成 Action 定义与文件 manifest。 | 独立 Word 包不夹带不再需要的 Session/managed Session 执行栈，迁移位置后可独立发现、提交和查询。 |

- [x] 按 C01–C13 逐项落实，并记录对应改动与验证位置。
- [x] 删除 Word 专属失效文件、导出和异常类型；共享文件先查实际调用者，不能按文件名批量删除。
- [x] 同时检查源码、安装包、示例和工程入口，避免仅 Skill 文案宣称完成迁移。

### 必须保留的执行保护

- 精确定位用户指定文档，不使用活动窗口、Selection、同名文件或打开顺序决定目标。
- acquisition guard → Lease 无解锁空隙；Save As 保留原文档绑定和必要的新旧路径/文件身份保护。
- Task 之间隔离；同时操作同一文档发生明确冲突，不自动排队、抢占或强制解锁。
- 效果不确定时不重放；进程退出不等于 WPS 已停止执行，不能无证据清除 Quarantine。
- 自动化进程受控且清理有界；WPS 应用与用户文档不是可随 Task 杀掉的资源。
- 参数和返回数据校验、操作原生回读、Revision/Range、原有未保存修改保护继续有效。
- 重复提交与回执查询只返回已有事实；丢失 stdout、清理失败或输入删除失败不授权重做文档动作。

## 4. 现有测试的处置清单

先整理现有测试的责任，再补确实缺失的断言。每个行为指定一个主要测试位置；跨层集成测试只验证连接后的实际行为，不完整复制单元测试矩阵。测试数量不作为完成指标。

下列路径相对 `src/test/python/tests/`。

| 现有测试或夹具 | 处置 | 应留下的价值 / 应移除的历史假设 |
| --- | --- | --- |
| `client/test_word_task.py` | **迁移、删旧分支** | 保留固定顺序、空 completion、原有修改保护、失败停止、部分效果、保存失败停止 PDF、引用解析。去掉 v2、手工 taskId、checks 执行和 stdin 成功路径。 |
| `client/test_word_submission.py` | **作为当前提交与回执主位置更新** | 保留自动身份、输入消费、查询、冲突、并发、崩溃窗口、无重放；移除 v3 字段和“真实 Session 协议”前提，换成真实 Task 执行入口。 |
| `client/test_task_client.py`、`client/task_cli_fixture.py` | **按应用和责任拆分、合并重复项** | Word 有效的异常/回执行为迁入当前 Word 测试；旧 checks、SessionClient、手工 ID 行为仅在 Excel/PPT 仍需时保留，并明确应用范围。 |
| `client/test_session_client.py`、`client/test_managed_session.py`、`host/test_session_host.py` 及 Session 夹具 | **保留 Excel/PPT 所需部分，移除 Word 旧入口正向覆盖** | 不再为了 Word 测试继续支持 Session；必要的旧入口拒绝断言归入 CLI 当前入口测试。 |
| `core/test_action_session.py` | **按责任迁移，保留仍使用 Session 的应用覆盖** | 将绑定、Lease、部分获取、并发清理、未知效果、资源所有权等适用 Word 的断言迁到 Task；通用类型/契约验证留在通用测试；不全量复制原文件。 |
| `core/test_trace_journal.py` | **按诊断所有者调整** | 保留追踪可靠性与诊断降级；Word 改为 Task/Action 关联，其他应用仍可验证 Session 诊断。 |
| `word/test_adapter.py`、`word/test_handlers.py`、`word/windows/test_backend.py`、`word/windows/test_resources.py` | **保留操作边界，更新组装方式** | 精确对象、请求/结果验证、后端异常分类继续保留；直接构造 ActionSession 的 Word 用例改为 Task，移除失效 Session 文件导入。 |
| `word/test_contracts.py`、`word/test_output_retry.py`、`word/test_line_spacing_verification.py` | **保留独立 Action 回归，必要时只改依赖** | 本轮没有撤销这些操作及验证规则，不因测试年代或 Session 重构而删除。不要再加一份同义 Task 测试替代这些验证。 |
| `windows/test_document_coordinator.py`、`test_owned_process.py`、`test_bridge_runtime.py`、`test_powershell_bridge.py` 等 | **保留通用保护测试** | 操作系统资源、隔离、超时、Quarantine、桥接机制仍有效；在 Task 集成层补交接缺口即可。 |
| `cli/test_call.py`、`cli/test_build_word_skill.py`、`cli/test_build_skill.py`、`cli/test_repository_scripts.py` | **更新现行入口与打包边界** | Word 旧入口必须拒绝；新包独立执行；Word 包不含无用旧栈；Excel/PPT 入口和打包不回归。 |
| `word/test_demo.py`、`word/live_acceptance.py`、`word/desktop_acceptance.pyw`、`word/write_content_acceptance.py` | **整合保留有用途的入口** | 迁移仍验证独特行为的场景，删除仅重复演示旧 Session 调用方式的脚本或片段。 |
| `persistence/word_task_v3_acceptance.py` | **迁移为无版本的 Word Task 原生验收主入口** | 更新请求、生命周期和查询断言，去掉文件名及说明中的 v3 假设；避免保留一份旧脚本再另加“新版”。 |
| `persistence/word_output_retry_acceptance.py`、`persistence/live_acceptance.py`、持久化 Session 夹具 | **保留独特机制，迁移 Word 分支** | 冲突候选、原有文件不变、Save As 后绑定和 Lease、原生输出验证继续有价值；Excel/PPT 分支继续使用其当前入口。 |
| `build/evidence/` 内历史脚本、样本及报告 | **保留为历史证据，退出当前执行入口** | 不复制为新一代测试框架，不把历史样本或历史通过数计入迁移后的验收结果。 |

- [x] 在实施时给被删除或移动的测试注明“行为撤销／已迁移到何处／与何处重复”。
- [x] 移除 Word 的旧协议成功测试，同时保留精简的拒绝覆盖，验证没有绕过路径。
- [x] 保留测试夹具的断言，不把验收断言重新放回 Agent 的 Task Request。

### 执行侧验收必须回答的问题

这些是行为边界，不要求每行创建一个新测试文件；优先调整对应现有测试。

| 验收问题 | 最小可证据化结果 |
| --- | --- |
| 当前请求能否独立执行？ | 无 `version` 请求通过正式入口；Word 旧协议与 Session 入口执行前拒绝；发现与查询不启动 WPS。 |
| 默认不保存是否真实成立？ | 新建不保存：内容留在打开的未保存文档；编辑已有文件不保存：内存修改可见、磁盘原文件哈希不变；只导出不隐式保存 Word。 |
| 固定顺序是否由 Task 保证？ | 获取一次；内容按计划；仅执行选择的交付；保存先于 PDF；非法计划在文档效果发生前拒绝。 |
| Action 成功是否有依据？ | 参数与结果契约、必要的原生回读仍有效；不能仅依靠 fake 返回 succeeded 就宣布操作功能通过。 |
| 中途失败是否保留事实？ | 已完成步骤保留，失败或 unknown 步骤明确，后续未执行；没有回滚、重建或自动重放。 |
| 保存失败是否真正结束？ | 受控失败下内容保留、保存异常可见、后续 PDF 未执行、资源清理、不跨 Action/Task 自动补救。Action 内既定的有限冲突候选机制单独验证。 |
| 绑定和 Lease 是否归 Task？ | 活动文档变化不转移目标；同文档竞争受拒；正常结束后安全释放；获取中断和可能仍在执行时保留正确保护。 |
| 旧范围是否仍受保护？ | 同 Task 修改后的旧范围拒绝；不同 Task 不复用范围，即使重新打开同一文件也要重新取得。 |
| 提交、崩溃和查询是否不重放？ | 相同输入查询、并发、输入已消费、回显丢失、持久化中断均不产生第二次文档效果。 |
| 清理与打包是否完整？ | 成功、失败、取消、进程丢失均有准确清理事实；独立 Word 包不依赖仓库或旧 Session；共享改动不破坏 Excel/PPT。 |

Agent 是否理解“用户没要求保存”、是否在保存失败后只报告而不补救，属于后续 Agent 验收；执行侧通过不能冒充这两项已验证。本轮清单不恢复 E01/E02/E04 九样本，也不进入稳定性长测。

## 5. Skill、领域定义与 ADR 迁移

| 位置 | 后续工作 |
| --- | --- |
| `src/main/resources/skills/wps-word/SKILL.md` | 使用无版本请求；说明 Task Client 提交/查询；新建和编辑默认不保存；空 completion；内容异常与保存异常的不同处理；停止、保留效果、无隐式清理保存；清除旧入口和 Session 操作指导。 |
| `references/actions.json` | 继续从生产注册生成。只有实际契约受影响才变更，不另手写一份 Task 验证契约或生命周期 Action。 |
| `CONTEXT.md` | 定义清楚 Word Task、Task Client、Task Executor、Task Request/Response；Word 绑定及 Content Revision 的生命周期归 Task；Word 的 Persistence Intent 默认不保存。Session 术语明确限于尚未迁移的 Excel/PPT，不能全库删掉仍有效的定义。 |
| ADR 0028 | 本轮新决策取代其“Word 内部保留 Session”“按 v3 请求执行”“保留 Word 低层旧入口”等选择。新 ADR 落定时标记 0028 已取代，保留历史正文。 |
| ADR 0001、0007、0009、0012 | 分清 Word 新 Task 规则与 Excel/PPT 尚有效的 Session 规则，涵盖生命周期、响应身份、绑定与 Lease。不能因 Word 迁移就将三应用共享决策全部作废。 |
| ADR 0002、0003、0005、0016 | 校准三应用分阶段迁移、共享机制、控制面和硬切换适用范围，清除“Word 仍统一走 Session”的现行指引。 |
| ADR 0014、0015、0017、0018 | 保留 Action 集合、验证与准入原则；将 Word 内容范围、Revision 和文档执行环境归属校正为 Task。 |
| ADR 0026 | 保留保存保护和 Action 内冲突处理；Word 的释放及失败终止归 Task；更新指向已取代 0028 的现行链接。 |
| ADR 0020 及 Excel/PPT 专属说明 | 保留其当前行为。仅在确有引用受影响时更新链接，不借此迁移其他应用。 |
| 当前决策入口（建议 `docs/adr/README.md`） | 列出仅对当前模型有效的 ADR、适用应用和主题。被取代 ADR 不混入当前必读清单。 |
| 设计演进记录（建议 `docs/design-evolution.md`） | 按时间串起原方案、问题或用户目标变化、替代决定、理由及对应 ADR。只有历史入口引用已取代记录。 |
| `README.md`、`README.zh-CN.md`、`INSTALL.md`、`scripts/README.md` | 同步 Word 当前流程，删除 Word 旧协议演示及“已有文件默认保存”说明；保留其他应用仍有效的流程。 |
| `WPS_WORD_TEST_PLAN.md`、`WPS_SKILL_DESIGN_NOTES.md`、`WPS_SKILL_ISSUES_20260916.md` | 明确哪些属于历史，更新当前接续点与现行决策链接。历史结果保留，不直接修改成新架构已经通过。 |

- [x] 用一份聚焦的新 ADR 记录本轮决策，说明仅迁移 Word、去 Session、去旧入口、无请求版本、默认不保存、保存失败终止的原因与代价；不为每个实现小步骤新增 ADR。
- [x] 对共享 ADR 做应用范围校准；部分仍有效的文档不能只贴整篇 superseded 而丢失 Excel/PPT 规则。
- [x] 当前正文和阅读入口不再把已取代 ADR 当现行依据；历史 ADR 的替代标记与演进链接保持可追溯。
- [x] 演进原因区分“用户重新明确目标”和“有实验证据”；不补写没有证据的因果收益。

演进记录至少应覆盖：逐 Action Session 调用 → 单次预编排 Task → 三部分 Action 结构 → task file 与持久回执 → 代码承担身份和验证、内部仍是 Session → 本轮 Word Task 完整接管资源并删除 Word 旧入口。每段保留真实的决定原因，历史仅描述其当时适用范围。

## 6. 实施顺序与完成门槛

1. **固定契约与文档边界。** 按已确认结论更新领域定义、ADR 与演进记录，给出无版本请求、空 completion 和保存失败响应的当前示例。
2. **先整理测试责任。** 按第 4 节标明哪些迁移、删除、保留及目标位置，再随实现改动测试；不先堆一批端到端样本。
3. **迁移 Word 执行侧。** 拆开必要的通用机制，将绑定、Lease、Action 调用及有界清理迁到 Task，消除 Word 对旧 Session 执行链的依赖。
4. **一次切换入口、回执和包。** 无版本请求、当前提交/查询、旧入口拒绝、演示与打包同时对齐，避免新入口仍跑旧栈。
5. **验证执行侧。** 按改动运行必要的契约、Task、资源和共享依赖回归，再用隔离 Windows 文档核对实际效果；出现缺口只补对应责任位置。
6. **核查清理结果。** 检查 Word 依赖、包内容、旧入口、重复测试和文档链接；确认 Excel/PPT 的保留边界没有被破坏，记录真实验证范围。

完成迁移至少同时满足：

- [x] Word 生产路径不再实例化 SessionClient、SessionHost、ActionSession，也不通过它们执行任何 Action。
- [x] Word 专属 Session 文件与公开工程入口已删除；共享 Session 只服务仍未迁移的应用。
- [x] 无版本请求、默认不保存、保存失败终止、资源释放均在正式 Task 路径成立。
- [ ] 关键保护行为有当前证据；有效测试已迁移，旧协议正向测试退出 Word 验收。
- [x] 安装包、脚本、Skill、CONTEXT、有效 ADR 与当前测试说明一致。
- [x] 历史证据独立标注；没有将旧样本数量或 Agent 交付结果当作新执行侧的通过依据。

## 7. 工作区与本轮产物

当前工作区已有大量未提交修改和历史测试产物。迁移实施时按职责逐文件处理，不 reset、clean、整体覆盖或自动提交；`decision.md` 是用户笔记，不修改。历史 evidence 保留，不为“干净”删除原始失败证据。

用户已于 2026-09-17 授权执行侧迁移。WorkBuddy Agent 样本测试仍不运行。实际完成与验证记录见下方。

主要核对入口：[当前 Task 执行器](../../../src/main/python/wps_skills/client/task_client.py)、[Word 计划编译](../../../src/main/python/wps_skills/word/task_plan.py)、[Word Task 组装](../../../src/main/python/wps_skills/word/windows/task.py)、[共享执行资源](../../../src/main/python/wps_skills/core/action_runtime.py)、[Skill 构建](../../../src/main/python/wps_skills/cli/build_skill.py)、[ADR 0029](../../adr/0029-make-word-task-the-execution-boundary.md)。这些链接已更新为当前实现；历史差距见第 2 节。

## 8. 实施记录（2026-09-17）

**本轮完成：** Word 执行侧、入口、打包、现有测试整理和领域文档已切换；本地回归通过。上文代码项的勾选只表示实现与本地证据，不能代替 Windows 原生验收。完整迁移门槛中的“关键保护行为有当前证据”保持未勾选，待真实 WPS 验证。

### C01–C13 实际改动与证据

以下源码路径相对 `src/main/python/wps_skills/`，测试相对 `src/test/python/tests/`。

| 编号 | 实际完成 | 当前验证位置 / 尚缺证据 |
| --- | --- | --- |
| C01 | `word/skill.py` 和安装脚本仅导出 Word Task CLI，限制 app word；移除 open_session 和 Session 异常导出。 | `cli/test_build_word_skill.py`、`test_build_skill.py`。 |
| C02 | `cli/call.py` 执行前拒绝 Word Session、managed、stdin、消费开关、debug 入口；生产 Session 工厂与 managed API 拒绝 Word。Excel/PPT 路由保留。 | `cli/test_call.py`；Excel/PPT Session/managed 回归。 |
| C03 | `word/task_plan.py` 直接生成无版本内部计划；拒绝 version/taskId/checks；校验 ID、前序引用、章节 Action 与参数；带引用的已知结构也在获取前预检。 | `client/test_word_task.py`、`test_word_submission.py`；完整参数仍在派发前校验。 |
| C04 | `client/task_client.py` 直接驱动 `word/task.py`；Excel/PPT 旧 Task 执行隔离到 `legacy_task_client.py`，Word 无 Session 子进程与握手。 | `client/word_task_cli_fixture.py` 使用真实 WordTask/资源机制并断言未导入旧栈；提交与进程丢失测试。 |
| C05 | 契约、验证、Runtime、绑定和资源机制抽到 `core/action_runtime.py`；`action_session.py` 仅为 Excel/PPT 构造 Session 响应，复用同一资源机制。 | Core/Windows 原有保护测试保留；Word 与共享 Windows 依赖不再导入 action_session。 |
| C06 | WordTask 拥有同一资源聚合；提交绑定/Lease 后才返回成功；部分获取、失败、取消、超时和结束统一释放；Adapter 清理失败单独记录且等待有界。 | `word/test_task_resources.py`、`windows/test_document_coordinator.py`；真实同文件竞争及在途 Quarantine 待 Windows 验证。 |
| C07 | 新增 `word/windows/task.py`，删除专属 `word/windows/session.py`；构造失败保留清理事实；Word Task 对 Action 有界等待；阻塞管道不阻止 launcher 终止和 Job 释放。 | `word/test_task_resources.py`、`windows/test_owned_process.py`；操作系统层实机验证待执行。 |
| C08 | Word Adapter/handlers/contracts/registry 改用公共类型；绑定错误使用 TASK 前缀；既有原生回读、参数/结果验证及 Action 内冲突候选规则保留。 | Word adapter/handlers/contracts/output_retry 回归；没有新增 Task 保存重试。 |
| C09 | 新增 `client/json_io.py`，保留严格 JSON、Windows 文件锁短暂重试和原子写入；当前回执按应用结构处理，版本仅用于辨认历史磁盘证据。 | 提交身份、消费、冲突、持久化窗口、查询、owner loss、历史回执测试；旧记录不删除、不重放。 |
| C10 | Task Response 保留获取/步骤/交付、停止与清理事实；Word trace 关联 taskId/traceId，无 Session 身份。启动释放异常、Action 超时、cleanup owner loss 不抹去已确认效果。 | `client/test_word_submission.py`、`word/test_task_resources.py`、Core trace 回归。 |
| C11 | Word bridge 删除 debug close/quit 分支；普通资源释放不保存或关闭文档。Revision 保留每次获取独立 GUID 与内容变更刷新；Windows 桌面 Session 名称保留。 | 资源/契约检查；跨 Task 旧范围已编入原生验收，尚未执行。 |
| C12 | Word demo 使用完整 Task；段落与输出重名验收迁为 task-file；`word_task_v3_acceptance.py` 移为 `word_task_acceptance.py`。混合 Session 验收仅留 Excel/PPT。 | 演示入口、脚本导入与编译检查；原生脚本断言未作为本轮实机通过证据。计划提到的 `word/desktop.py` 原本不存在。 |
| C13 | Word 构建排除 host、Session Client、managed、legacy Task 和 ActionSession 模块；重新生成 actions.json 与 manifest。 | 三应用独立包导入/发现；Word 包迁移后拒绝旧入口、提交至平台边界并查询回执；最终包 41 文件哈希核对。 |

### 测试处置台账

- `client/test_word_task.py` 保留固定顺序、空 completion、原有修改保护、失败停止、保存失败停止 PDF、引用；撤销的 3 个 checks 成功场景和 stdin 成功场景删除，拒绝覆盖集中到 submission/CLI。补取消和带动态引用的静态非法字段预检。
- `client/test_word_submission.py` 改为无版本正式提交；原 Session 协议夹具换为真实 WordTask 夹具；保留输入、并发、查询、消费与冲突断言，补执行/清理阶段进程退出和历史回执只读查询。
- `client/test_task_client.py`、`task_cli_fixture.py`、Session Client/managed/Host 测试明确只承担 Excel/PPT 的旧协议行为，不再正向测试 Word Session。
- Core 绑定矩阵不复制：Session 测试使用 Excel 假适配器，公共假资源抽到 `core/resource_fixture.py`；WordTask 只补连接后的获取、绑定、未知效果、清理、超时和诊断断言。Adapter/Backend 直接构造 WordTask，原断言保留。
- `test_contracts.py`、`test_output_retry.py`、`test_line_spacing_verification.py` 保留独立 Action 责任；本轮不以 fake 成功替代原生效果验证。
- 原生 Word Task 验收文件改名而非复制；新增空 completion、内存修改/磁盘哈希不变、仅 PDF、保存失败停止 PDF、跨 Task 旧范围拒绝场景。段落验收保留格式、插入范围、空段落、inline、表格与持久 XML 检查。
- 混合持久化旧 Word Session 分支移出当前入口；同一 Task 多次保存/失败后继续等已撤销编排不保留兼容。保存引用不变与协调机制继续由 persistence Backend、Coordinator、Action 测试承担；原生 Lease/Quarantine 复核仍是待办。
- 历史 `build/evidence/` 不改写，不把旧 Agent 样本数量计入本轮。`decision.md` 未修改。

### 实际验证与产物

- 本地环境：macOS（Darwin），无 Windows PowerShell/WPS。本轮未操作 Windows，未运行任何 WorkBuddy Agent 样本或稳定性长测。
- 全部本地单元/集成回归：`PYTHONPATH=src/main/python:src/test/python python3 -m unittest discover -s src/test/python/tests -p 'test_*.py'`：**306 项，305 通过、1 跳过**。跳过的是需要 Windows PowerShell 的行距验证。
- `python3 -m compileall -q src/main/python src/test/python scripts` 和 `git diff --check` 通过。
- 最终独立包：[wps-word](../../../build/archive/runs/word-task-migration-20260917-final/wps-word)，41 文件 manifest 与源码核对通过。未覆盖已有构建目录，未安装到 WorkBuddy。
- 当次回归日志：[执行侧验证](../../../build/evidence/word-task-migration-20260917/verification.txt)。本地 mock/夹具测试证明执行连接与保护逻辑，不证明 WPS COM 实际效果。
- `docs/`、`CONTEXT.md`、`FILE_STRUCTURE.md` 原受仓库 `.gitignore` 忽略；文档改动已写入工作区，未改变忽略规则或暂存文件。未 reset、clean 或自动提交；Excel/PPT 的已有未提交 Skill 修改保留。

### 尚未完成的原生验收

- [ ] 在隔离 Windows 桌面运行 `persistence/word_task_acceptance.py --skill <最终Word包> --output <新目录>`，核对无保存、仅导出、保存失败及跨 Task 范围保护。
- [ ] 运行迁移后的输出冲突、段落和 Word demo 原生验收，保留原生回读及持久文件证据。
- [ ] 核对真实同文档竞争、活动文档切换、Save As 后新旧身份保护、在途进程丢失/Quarantine 与有界回收。现有通用机制测试不是这组实机证据。
- [ ] 原生证据齐备后再勾选第 6 节最后的保护验收门槛；Agent 验收仍须另行安排，本轮不执行。


## 2026-09-17：Word 文件职责整理

本节记录迁移后的结构调整；上文路径保留为执行当时的历史记录。现行文件职责见 [FILE_STRUCTURE.md](../../archive/FILE_STRUCTURE-20260917.md#生产-pythonword-应用)。不改变已确认 Task 设计。

- [x] `word/task.py`、`task_plan.py` → `task/executor.py`、`task/plan.py`。
- [x] `adapter.py` 拆成 `actions/adapter.py` 与 `backend/` 下的协议、operation、preparation、acquisition、异常模块；handlers、registry 归入 `actions/`。
- [x] `contracts.py` 拆成 `contracts/` 下的 definitions、schemas、formats、validation、examples；包入口保留正式能力发现。
- [x] `output_names.py` → `persistence/output_names.py`；Windows 组装入口 → `windows/task_factory.py`。
- [x] 既有测试按职责迁至 `tests/word/{actions,contracts,persistence,task}`，更新调用方、子进程夹具和打包导入。
- [x] 本地回归：306 项，305 通过，1 项 Windows PowerShell 环境跳过。
- [x] 独立 Word 包：所有模块隔离导入通过、资源路径有效、清单校验通过；14 个 Action 的 `actions.json` 与结构调整前逐字节一致。产物为 `build/archive/runs/word-structure-20260917/wps-word`，证据为 `build/evidence/word-structure-20260917/`。

未运行 Windows 原生验收或 WorkBuddy Agent 样本测试；未修改 Excel/PPT 执行结构、未自动提交。


## 2026-09-17 后续：正式替换与定向验收完成

已将观察时的页眉页脚 Range 读取替换为 live Document.WordOpenXML，并同步 contract 约束、生成的 actions.json、资源打包和测试。Windows 原生指纹 4 例、正式 Task 5 例通过；读取保持保存状态，未保存页眉编辑能改变指纹。上文“尚未替换”等描述保留为此前阶段记录；当前结果、包和证据见 [指纹修复验证](WPS_WORD_FINGERPRINT_FIX_RESULTS.md)。此次不代表原 18 项清单或迁移计划全部验收完成。


后续验收更新（2026-09-17 22:56）：原 18 项功能清单已全量重跑，18 个 Task、80 个 Action 响应全部满足预期；人工粗审待进行。见 [最新实机结果](WPS_WORD_FUNCTIONAL_TEST_RESULTS.md)。上文未重跑的描述为当时阶段状态。
