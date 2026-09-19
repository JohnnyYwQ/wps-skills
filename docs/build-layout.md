# 构建产物与证据目录

这里存放构建产物与运行证据，不是源代码目录。2026-09-19 已归类；当前安装包由当时未提交工作区重建，不等于已发布的 GitHub Release。

| 目录 | 内容 | 用途 |
| --- | --- | --- |
| `plugins/wps-skills/` | 兼容 Codex/Claude Code 的统一插件分发目录；同级有 ZIP 和 SHA-256 校验文件 | 在 Windows 通过宿主的插件命令安装 |
| `skills/` | 当前完整的 wps-word、wps-excel、wps-ppt | 将整个应用目录安装给 Agent |
| `tools/wps-doctor/` | 独立环境诊断及 Excel 实际验收工具 | Windows 执行 `python doctor.py --app excel` |
| `tools/action-acceptance-v6/` | Word/Excel/PPT 全 Action 独立测试包 | 已在 Windows 通过固定 JSON 验收 |
| `tools/design-value-v13/` | 设计收益实验工具，历史批次各有独立版本 | 运行固定 Task 与受控对照；不是宿主安装包 |
| `distributions/` | 三份 Skill、doctor、Action 验收和设计实验 ZIP 与 `manifest.json` | 复制分发；清单记录 SHA-256 和构建验证 |
| `runs/startup-communication/` | 独立运行包、远端元数据、快照 | 用调试入口查询或读取，不是正式安装包 |
| `runs/action-acceptance/` | 三应用测试请求、响应、产物和失败记录 | 最终完整通过轮次为 v6 |
| `runs/design-value/` | 不依赖宿主的机制、故障与成本实验 | 原始请求、响应、文档观察和哈希归档 |
| `runs/doctor/` | 后续 doctor 默认输出位置 | 每次一个新目录 |
| `evidence/` | 历史验收证据与回归日志 | 保留可追溯记录 |
| `archive/packages/` | 旧安装包、发布材料 | 不作为最新版安装 |
| `archive/runs/` | 旧修复轮次、排障材料及测试输入输出 | 同轮文件整体保存，未拆散或改写 |
| `archive/loose-files/` | 旧散落 ZIP、日志、脚本、说明 | 历史材料 |
| `worktrees/` | Git 管理的发布分支工作区 | 不是安装包，不能按普通临时文件删除 |

## 现在该拿哪一份

- 统一插件：`plugins/wps-skills.zip`，解压后按根目录 `README.md` 安装，三个 Skill 一并提供。
- Word：`distributions/wps-word.zip`，解压得到完整 `wps-word/`。
- Excel：`distributions/wps-excel.zip`，解压得到完整 `wps-excel/`。
- PPT：`distributions/wps-ppt.zip`，解压得到完整 `wps-ppt/`。
- 环境诊断：将 `distributions/wps-doctor.zip` 解压到独立目录，运行里面的 `doctor.py`；基础检查支持三应用，真实文档验收目前支持 Excel。

- 全 Action 批量验收：`distributions/wps-action-acceptance.zip`，解压后运行 `run_word.py`、`run_excel.py`、`run_ppt.py`。
- 设计收益实验：`distributions/wps-design-value.zip`，实际运行过的 v13 包；解压后运行 `run.py` 并选择实验组。完整结论见 `docs/testing/host-independent-results.md`，各批次使用版本见 `build/evidence/design-value/package-catalog.json`。

每个 Skill 自带共享运行时，可单独安装，不含 demo。doctor、Action 验收和设计实验工具单独分发，不是 Application Skill；测试包内自带待测 Skill，不替换宿主安装包。

## 更新规则

修改 `src/` 后重新构建，不直接编辑安装包。统一插件入口为 `scripts/build/plugin.py`，默认输出到 `build/plugins/wps-skills/` 并生成同级 ZIP。单应用入口仍是 `scripts/build/{word,excel,ppt}.py`；默认输出到 `build/skills/wps-<app>`。已有目录会拒绝覆盖。构建新版本时指定新输出目录，验证后归档旧包再更新当前交付目录及 ZIP、清单，避免混合版本。

调试默认输出到 `runs/startup-communication/`；仓库 doctor 默认输出到 `runs/doctor/`。历史材料中的旧绝对路径保留作为当时证据，不能保证其中临时脚本迁移后仍可直接重跑。旧位置查 `archive/path-map-20260919.json`；内容核对查 `archive/move-verification-20260919.json`。

本目录被 Git 忽略，但 **不能把整个 build 当作随时可删除的缓存**：其中有尚未外部备份的实机证据和注册的 Git worktree。源代码、固定测试资源及维护中的文档分别位于 `src/`、`scripts/`、`docs/`。当前本地目录入口见 [build/README.md](../build/README.md)。
