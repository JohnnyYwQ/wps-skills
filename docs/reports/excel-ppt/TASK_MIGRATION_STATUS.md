# Task 迁移状态

Excel/PPT 已接入与 Word 共用的 Task 执行机制，当前业务代码位于 `src/main/python/wps_skills/{excel,ppt}`，原生操作位于 `src/main/resources/wps_skills/{excel,ppt}`。不再保留另一份迁移源码或旧 Session 执行入口。

构建与验证入口见 [脚本入口](../../../scripts/README.md)，本轮实机范围和结果见[Task 迁移验证报告](WPS_EXCEL_PPT_TASK_MIGRATION_RESULTS.md)。历史 Session 证据不计作当前 Task 验收。
