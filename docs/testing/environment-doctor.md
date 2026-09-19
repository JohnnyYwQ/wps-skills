# Windows 环境诊断

在目标 Windows 上运行。Python 3.10+，标准库，无第三方依赖。此工具先做环境验收，不会在每次正式 Task 前自动运行，也不会自动安装软件、修复注册表或调整权限。

## 仓库内使用

```sh
python scripts/doctor.py --app excel
python scripts/doctor.py --app word
python scripts/doctor.py --app ppt
python scripts/doctor.py --app excel --smoke-test
```

默认检查不创建文档、不激活 WPS：检查 Windows、Python、输出目录读写、系统 Windows PowerShell 5.1、用户桌面会话、当前用户可见的对应 COM 注册，以及 UTF-8 stdin/stdout 往返（中文、空格、引号、反斜线、换行、emoji）。报告还记录 PowerShell 位数、STA/MTA 和是否管理员。权限与位数是观察值，不凭管理员或 64 位直接宣称兼容。

成功退出码为 0，有失败为 1；参数错误为 2。终端列出逐项状态及失败修复建议，完整数据写入 `report.json`。默认结果存于新的 `build/runs/doctor/doctor-时间-随机值/`。可用 `--output 路径` 指定**不存在**的目录，禁止重复使用旧目录。

`--smoke-test` 当前只支持 Excel。基础检查通过后，复用已验证的 Task 测试：创建文档 → 写入/回读/保存 → 独立 XLSX 检查 → 关闭测试文档；再复制产物，重新打开修改并验证。失败立即停止，保留证据和仍存在的文档现场。Word/PPT 的基础环境检查可用，真实文档验收尚未接入；请求其 smoke test 会明确失败，不能当作已验证。

默认检查通过时 `documentOperationsVerified=false`；仅真实文档验收通过才为 true。注册存在不等于能够连接；本机通过不代表其他账户、其他 WPS 版本或长时间运行必然可靠。

## 给另一台 Windows 的独立包

在开发机生成隔离包：

```sh
python3 scripts/debug/startup_communication.py pack --config src/test/resources/diagnostics/startup-communication/com.json
```

复制打印出的运行目录下 `bundle.zip`，在 Windows 解压到新目录。进入解压目录运行：

```sh
python doctor.py --app excel
python doctor.py --app excel --smoke-test
```

包内包含测试所需的 Excel Skill、原生资源、配置及哈希清单。目标机器不需要 Git、项目源码或额外 Python 库。默认报告在解压目录的同级新目录；不得把输出写进包内，避免修改哈希或递归复制自身。运行实际验收时先校验包，再复制到该次证据目录中独立执行。

若 Python 尚未安装，脚本本身无法启动，需要先安装解释器；若指定输出目录不可写，工具只能在终端输出失败，无法保证在那里保存报告。

## Mac 调度与 Windows 本机环境分开

SSH 只是远程调度手段，不是产品本机运行前提。直接通过 SSH 运行 doctor 时，Session 0 会报桌面检查失败，其余可检查项仍会给出结果。应在 Windows 已登录桌面运行，或通过已有交互式计划任务调度。安装 SSH、配置账号、上传包和注册计划任务属于远程调度检查，继续使用 [启动通信工具](startup-communication/README.md)，不混入 WPS 本机依赖。

PowerShell 基础探针设有 15 秒超时，超时后终止并等待这个直接拥有的进程；它不启动 WPS 或其他子进程。实际文档测试沿用用例看门狗和 bridge Job Object 管理。不能把 Mac SSH 命令超时误认为 Windows 已退出。

## 2026-09-19 验证

独立包运行标识：`startup-20260919T103327-5a3c9bd5`。

- Session 0：依赖、注册、UTF-8 通信通过；桌面检查失败，未执行文档操作。
- Session 3：基础检查全部通过；Excel 新建及重新打开 2/2 通过，均走 COM 激活路径；独立文件验证和测试文档清理通过。
- 环境：Windows 11，PowerShell 5.1.26100.9444，64 位、非管理员、STA；实际 Python 自报 3.13.14（解释器所在目录名不作为版本依据）。

本地证据：`build/runs/startup-communication/startup-20260919T103327-5a3c9bd5/doctor-desktop-evidence/` 与 `doctor-ssh-evidence/`。只验证了一台 Windows；Word/PPT 基础检查的选择逻辑已有实现，本轮实机使用 Excel。
