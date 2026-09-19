# Action 执行端验收计划与 Windows 批量工具

目标：先排除 Agent 选择和填参的不确定性，用固定 Task Request JSON，验证全部正式 Action 的可执行性、Action Response 和 Task Response。执行端基线通过之后，再进行设计选择实验。本工具是测试包，不是 Application Skill。

## Windows 手动运行

将独立 ZIP 解压到新目录，在已登录 Windows 桌面的普通终端执行：

```powershell
python run_word.py --root C:/wps-tests/word-新的轮次
python run_excel.py --root C:/wps-tests/excel-新的轮次
python run_ppt.py --root C:/wps-tests/ppt-新的轮次
```

Python 3.10+，系统 Windows PowerShell、对应 WPS COM 注册。测试包包含三份本轮源码构建的独立 Skill，无需安装 Git 或 pip 包；也不替换宿主安装的 Skill。每个 root 必须不存在，输出父目录由测试工具准备。无参数时在测试包 `runs/` 下自动生成唯一目录。

若需要先检查 JSON：

```powershell
python run_excel.py --root C:/wps-tests/excel-review --prepare-only
python run_excel.py --root C:/wps-tests/excel-review --run-prepared
```

请求和预期在执行前已经写入磁盘并记录哈希。修改方案后须生成新轮次；不在已执行 Task 上恢复或重放。参数 `--skill` 可选另一份完整的待测安装包，但契约和用例不匹配时必须报错，不能静默降低覆盖。

仓库对应入口：`scripts/acceptance/word.py`、`excel.py`、`ppt.py`。独立包构建：

```sh
python3 scripts/build/acceptance.py --output build/tools/action-acceptance-新版本
```

Mac 可用 `--prepare-only --windows-root C:/实际目标路径 --root 本地准备目录` 导出计划。上传后只能在声明的 Windows root 执行；搬动已经编入绝对路径的计划需要重新生成。

## 基线覆盖

| 应用 | 正式 Action | 执行场景 Task（含前置准备） | 每 Action 非法参数拒绝用例 |
| --- | ---: | ---: | ---: |
| Word | 14 | 3 | 14 |
| Excel | 34 | 14 | 34 |
| PPT | 37 | 14 | 37 |

清单来自待测包的生产契约。每个正式 Action 都必须有预期成功的调用，以及至少一个明确的 data 字段断言。名单变化、缺覆盖、用例参数不合法会在准备阶段报错。全 Action 覆盖不等于穷尽每个参数组合、所有边界值或所有异常环境。

- Word：文字写入、查找、范围替换、读取、新建/打开、表格、图片、分页、页面设置、页眉页脚、save/saveAs/PDF。
- Excel：多表统计、公式链和错误单元格、混合类型与字面文本、格式、排序筛选、复制清理、工作表与行列结构、合并、查找替换、行列尺寸、打开保存导出、过期 token 停止。
- PPT：多页内容、字体颜色、形状外观、对齐分布、顺序与几何、复制移动删除页、图片表格、备注设置、查找替换、打开保存及 PDF/PNG、过期 token 停止。
- 每 Action 的非法参数用例在真实 CLI 提交，预期 Task 在文档操作前拒绝，输入保留且无受理身份。预期拒绝不是实验失败。

## 断言与证据

1. Task：type、app、state、outcome、stop 阶段/步骤/错误码、cleanup、Task 身份、输入消费状态、回执查询一致性、退出码。
2. Action：执行顺序、address、state、response.outcome、Task/trace 身份、错误码、data 精确值/容差/长度；另用原参数和解析后的结果引用验证正式结果契约及语义约束。
3. 产物：真实文件、大小、哈希、OOXML 完整性和主 XML、PDF/PNG 格式签名；Word 增加指定文字、表格和嵌图的独立文件检查。Excel/PPT 多数业务效果由逐字段响应断言验证；容器检查不能冒充全部业务内容独立验证。
4. 人工视觉：报告始终标记 pending，不因自动断言通过而声称页面视觉效果合格。

```text
run/
  manifest.json         计划、Action 覆盖矩阵、请求/预期与包哈希
  requests/             保留不消费的原始 Task Request
  expected/             每一步和整个 Task 的预期
  submissions/          真正提交的副本；受理后可被执行器消费
  results/<case>/        stdout、stderr、Task Response、断言差异、回执和产物记录
  receipts/、traces/    正式执行器的回执与跟踪日志
  outputs/、assets/     测试产物与固定夹具
  report.json           实际进度、首次失败、实际 Action 覆盖
  run.claim             本轮已开始；禁止再次执行
```

意外失败立即停止当前批次，不重试、不悄悄放宽断言；未执行用例保持未执行。原生明确预期的失败按断言判定。失败现场保留，特别是发生过部分修改的文档。成功时仅在请求、回执和产物核对后尝试关闭精确路径、仍处于已保存状态的测试文档；不会 Quit WPS，也不会关闭其他文档。未能附加应用时记录保留，不冒充已关闭。

## 后续补测与验收门槛

第一道门槛：85 个 Action 的上述基线调用、响应断言和非法参数拒绝全部通过，所有非预期失败已解释且在新轮次复测；不能把失败前成功的部分称为整套通过。

第二道门槛按契约分支逐步加深：Word 的浮动图/分节/多节页眉、各应用边界尺寸、只读文档、文件冲突、格式/语言差异、更多独立文件内容检查。未覆盖分支明确登记，不能用 85/85 Action 名称覆盖率承诺“所有 Action 没有问题”。

Task 的进程终止、回执丢失、Lease/Quarantine 故障注入作为单独机制矩阵，不与正常 Action 成功率混算。Agent 测试使用执行基线已验证的范围，并把宿主失败单列归因。
