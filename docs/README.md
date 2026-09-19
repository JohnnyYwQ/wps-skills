# 项目目录导航

项目沿用 Java/Maven 风格的 source sets，语言仍是 Python 和 PowerShell。按职责放置文件；生产包的构建只读取 `src/main`。

```text
wps-skills/
├── README.md / INSTALL.md       使用与安装入口
├── AGENTS.md / CONTEXT.md       Agent 约定与领域词汇
├── src/
│   ├── main/
│   │   ├── python/wps_skills/   生产 Python 包
│   │   │   ├── core/           共享 Task 执行机制
│   │   │   ├── client/         Task 提交、文件与回执
│   │   │   ├── cli/            命令行与独立包构建
│   │   │   ├── windows/        Windows 进程与桥接
│   │   │   └── word|excel|ppt/ 各应用契约、执行和验证
│   │   └── resources/          Skill 模板、schema、原生脚本
│   └── test/
│   │   ├── python/tests/       回归测试与实机验收驱动
│   │   ├── python/diagnostics/ 调试工具实现（不进入生产包）
│   │   └── resources/          测试夹具、PowerShell、实验配置
├── scripts/                    薄命令入口：build/demo/validate/debug
├── docs/
│   ├── adr/                    架构决策（本地资料）
│   ├── agents/                 Agent 工作约定（本地资料）
│   ├── reports/                按 word、excel-ppt、shared 归档计划与报告
│   ├── testing/                调试工具使用说明和实验结果
│   ├── notes/                  设计笔记
│   ├── interviews/             本地访谈资料
│   └── archive/                过期结构快照等本地历史资料
├── experiments/                尚未纳入运行时的独立设计实验
└── build/                      构建产物、运行证据、临时文件（忽略）
```

## 日常入口

- 全 Action 执行端：[验收计划与批量脚本](testing/action-execution-plan.md)。
- 三应用 85 个 Action：[Windows 实测结果](testing/action-execution-results.md)。
- Agent 与设计选择：[价值验证计划](testing/design-value-plan.md)。

- 当前安装包、分发 ZIP 和历史证据：[构建目录说明](build-layout.md)。

- Windows 环境检查与独立包：[doctor 使用说明](testing/environment-doctor.md)。

- 构建、运行回归及实机验收：[脚本说明](../scripts/README.md)。
- 启动、通信与 COM 分段调试：[使用说明](testing/startup-communication/README.md)。
- 最近的 COM 两条连接路径：[验收结果](testing/startup-communication/COM_RESULTS.md)。
- Excel/PPT Task 迁移：[验证报告](reports/excel-ppt/WPS_EXCEL_PPT_TASK_MIGRATION_RESULTS.md)。

## 放置规则

一个 Module 的实现按职责聚集，Interface 通过既有 Python 包或薄脚本入口暴露。业务契约继续由各应用拥有；不为模仿 Java 而加入空的 controller/service/repository 层。

新测试放到 `src/test/python/tests/`，对应资源放到 `src/test/resources/`。调试实现与它的回归测试分开；Windows 测试包由构建器组装，源目录不必照搬部署布局。计划与实测报告保留各自明确的文件名，归入同一主题目录，避免拆散排障上下文。

`experiments/word-action-split` 是独立的候选契约设计实验，连同冻结 baseline 和 Skill 样本整体保留；不混入正式应用契约。已完成的迁移说明归入 [Task 迁移状态](reports/excel-ppt/TASK_MIGRATION_STATUS.md)，空的 `src/migration` 已移除。既有 `build/` 运行包和哈希证据保持原样，历史记录中的旧命令代表当时的运行方式。

本次路径迁移：根目录 `WPS_*.md` → `docs/reports/`；访谈资料 → `docs/interviews/`；旧 `FILE_STRUCTURE.md` → `docs/archive/FILE_STRUCTURE-20260917.md`；`outputs/` → `build/archive/runs/outputs/`；启动通信调试从 `experiments/startup-communication/` 拆入测试 source set。

根目录 `CONTEXT.md` 与 `AGENTS.md` 保持原位，便于工具发现。既有私有文档仍忽略；迁入的项目报告和调试说明可由 Git 跟踪。没有新增 Maven/Gradle 依赖或 Java 编译步骤。
