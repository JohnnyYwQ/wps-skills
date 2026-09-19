# 复杂业务任务实机测试 R01：准备阶段停止

2026-09-19。计划 Word / Excel / PPT 各 20 个业务 Task。Windows 交互桌面启动于 UTC 14:24:14，启动器于 14:24:18 以 exitCode 1 退出；没有自动重试。

## 实际结果

- 业务 Task：执行 0/60，未执行 60；不能报告业务成功率或业务性能。
- 素材准备：仅执行 `CW06-setup`。文档获取、文字写入及读取成功，保存失败，Task 为 `stopped/failed`。
- 保存错误：`OUTPUT_PARENT_NOT_FOUND`。本地准备生成了空 `outputs` 目录，但 ZIP 只收录文件，Windows 解压后缺失该目录。这是新测试包准备缺陷，不作为被测产品的保存功能回归。
- 同时存在字体观察断言差异：`seed_read` 前两段返回的 `westernFontFamily` / `eastAsiaFontFamily` 为 null，而预期是 Arial / 宋体；末段返回符合预期。保留为待定位项，未放宽断言，未据此宣布字体写入正确或错误。
- Task 清理结果为 succeeded，owned bridge 已释放。失败的未保存测试文档未被强制关闭，现场保留。

## 版本与后续改动

原始实测包为 `build/tools/complex-business-20260919-r01.zip`，SHA-256：

`604ea2cfb409c833ed4a9c9cce8ec51f6c48a90a1632770beaffb93b2601771a`

完整 60 例 JSON 与 expected 在启动前生成并记录哈希。业务 Action 计划次数 Word 162、Excel 419、PPT 741，合计 1322；14 个 setup Task 不包含在这些数字内。

本地契约/断言字段路径/覆盖检查通过，原有 18 项应用测试通过，新增 60 例准备校验通过。这些均不能替代 Windows 成功结果。

已在工作区测试入口补充执行前创建本轮专属 `outputs` 目录，并保留禁止重放检查。该修改尚未打包到新轮次或实机验证；R01 包和证据未改写。字体观察差异仍待定位。本轮按失败即停止约定结束。

## 证据

- [本轮汇总](../../build/runs/complex-business-20260919-r01/summary.json)
- [原始 Task Response](../../build/runs/complex-business-20260919-r01/runs/word/results/CW06-setup/response.json)
- [断言差异](../../build/runs/complex-business-20260919-r01/runs/word/results/CW06-setup/assertions.json)
- [启动器终态](../../build/runs/complex-business-20260919-r01/launcher.json)
- [归档文件哈希](../../build/runs/complex-business-20260919-r01/evidence-sha256.json)
- [60 例计划书](complex-task-execution-plan.md)
