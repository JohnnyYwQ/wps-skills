# 仓库脚本入口

以下命令从仓库根目录运行。所有 Python 入口支持 `--help`，也可用绝对脚本路径从其他工作目录启动。生产执行端需要 Windows 交互桌面；构建、查询、请求准备和离线汇总可在 macOS 上运行。

## 插件与独立 Skill

```sh
python scripts/build/plugin.py
python scripts/build/word.py --output build/skills/wps-word
python scripts/build/excel.py --output build/skills/wps-excel
python scripts/build/ppt.py --output build/skills/wps-ppt
```

统一插件输出到 `build/plugins/wps-skills/`，同级有 ZIP 和 SHA-256 校验文件，安装方法见 [INSTALL.md](../INSTALL.md)。已有输出不覆盖，再次构建请指定新的 `--output`。

三个 Skill 各自携带生产依赖。构建时按正式契约生成分文件 schema；开发源码中的全量 `references/actions.json` 不进入安装包。刷新源码中的定义使用 `python scripts/build/word.py --refresh-actions`，Excel/PPT 使用对应构建入口。

## 查询与提交

仓库入口是 `scripts/call.py`：

```sh
python scripts/call.py --app word --index
python scripts/call.py --app word --resolve writeContent saveAs
python scripts/call.py --app word --task-file C:/Tasks/request-001.json
python scripts/call.py --app word --task-status-file C:/Tasks/request-001.json
```

将 `--app` 改为 `excel` 或 `ppt` 可使用对应应用。请求由宿主文件工具写入 UTF-8 文件，每次新请求使用新路径；提交后的原路径用于查询回执。

安装包中的入口不同：`<skill-dir>/scripts/{word,excel,ppt}.py` 用于提交和查询回执，`<skill-dir>/scripts/schema.py` 用于按需查询定义。插件里的 `<skill-dir>` 位于 `plugins/wps-skills/skills/wps-<app>/`。仓库根目录没有 `scripts/word.py`。

## 环境检查与 Word 演示

```text
python scripts/doctor.py --app excel
python scripts/doctor.py --app excel --smoke-test
python scripts/demo/word.py --output-dir build/word-demo-new
python scripts/validate/word.py --output-dir build/word-acceptance-new
```

默认 doctor 只检查前置条件，支持三应用；`--smoke-test` 当前仅支持 Excel。Word demo 创建并保存一份演示文档，`validate/word.py` 检查这条演示流程，二者均需 Windows。`scripts/demo/word.ps1` 是同一演示的 PowerShell 启动入口，需指定 `-PythonPath`。输出目录使用新路径。

## 批量执行端脚本

| 入口 | 用途 |
| --- | --- |
| `scripts/acceptance/{word,excel,ppt}.py` | 对完整 Skill 运行相应应用的全 Action 验收；可先用 `--prepare-only` 生成 JSON 和预期响应。 |
| `scripts/build/acceptance.py --output <新目录>` | 构建包含 `run_word.py`、`run_excel.py`、`run_ppt.py` 的独立全 Action 包。 |
| `scripts/build/complex_acceptance.py --output <新目录>` | 构建三应用各 20 个复杂业务 Task 的独立包，同样只有三个应用运行入口。 |
| `scripts/acceptance/summarize_complex.py --root <三应用结果父目录> --output <汇总目录>` | 从现有回执和 trace 生成 JSON、CSV 与源文件哈希，不执行 WPS。 |
| `scripts/validate/applications.py --root <新目录>` | Excel/PPT 的原生 Task 基础流程检查；目录中需预先放入 `skills/wps-word`、`skills/wps-excel`、`skills/wps-ppt`。 |

独立包的三个脚本在 Windows 已登录桌面依次运行。固定请求、响应断言和运行边界见 [全 Action 脚本用法](../docs/testing/action-execution-plan.md) 与 [复杂业务脚本用法](../docs/testing/complex-task-execution-plan.md)。

## 分段调试

- `scripts/debug/startup_communication.py pack/start/status/collect/stop/report`：准备独立包、远程启动、查询状态及取回证据，见 [启动通信工具说明](../docs/testing/startup-communication/README.md)。
- `scripts/debug/design_value.py pack/start/status/resources/collect/report`：设计验证工具，见 [终端用法](../docs/testing/design-experiment-usage.md)。当前远程启动入口依赖 SSH 别名 `win` 和既有 Windows Python 路径；换机器需调整配置或使用包内 Windows 手工入口。

这些工具使用 `src/test/` 中的实现和资源，不依赖本地 `experiments/`，也不进入插件安装包。

## 本地检查与 CI

```sh
PYTHONPATH=src/main/python PYTHONUTF8=1 python -m unittest discover -s src/test/python -p 'test_*.py'
```

CI 使用 Linux/Windows × Python 3.10/3.13，运行源码编译、所有 Python 脚本入口检查、现行回归以及独立包搬迁检查；Windows 还使用 Windows PowerShell 5.1 解析原生脚本。全部通过后构建插件 ZIP，并保留为 Actions artifact。

CI 不启动真实 WPS。需要桌面和文档操作的脚本通过上述 Windows 手工入口运行。
