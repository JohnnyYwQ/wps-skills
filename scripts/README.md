# 仓库脚本入口

从仓库根目录运行下列命令。`scripts/` 只负责定位源码并传递参数；实际实现位于 `src/main/python`、`src/main/resources` 和 `src/test/python`。独立 Skill 安装后使用包内 `scripts/{word,excel,ppt}.py`，不依赖本目录。

```text
scripts/
├── call.py                 能力发现与 Session Host
├── build/
│   ├── word.py
│   ├── excel.py
│   └── ppt.py
├── demo/
│   ├── word.py / word.ps1
│   ├── excel.py / excel.ps1
│   └── ppt.py / ppt.ps1
└── validate/
    ├── word.py
    ├── excel.py / excel_common.py
    └── ppt.py / ppt_common.py
```

## 构建与发现

构建和能力发现可在 macOS、Linux 或 Windows 执行，不启动 WPS：

```sh
python scripts/build/word.py --output build/skills/wps-word
python scripts/build/excel.py --output build/skills/wps-excel
python scripts/build/ppt.py --output build/skills/wps-ppt
python scripts/call.py --app word --index
python scripts/call.py --app ppt --resolve openPresentation listSlides save
```

构建不覆盖已有目录；再次构建时使用新的输出位置。Session Host 使用 `python scripts/call.py --app <应用> --session`，其标准输入输出属于 JSONL Session Protocol。

## 可见演示

在装有 WPS 的 Windows 桌面普通 PowerShell 中运行。`.ps1` 入口检查 COM 注册并隐藏辅助进程的控制台，进度显示在当前终端：

```powershell
$python = (Get-Command python.exe).Source
./scripts/demo/word.ps1 -PythonPath $python
./scripts/demo/excel.ps1 -PythonPath $python
./scripts/demo/ppt.ps1 -PythonPath $python
```

三个演示都支持 `-OutputDirectory <尚不存在的目录>` 和 `-Delay <0到10秒>`；Python 入口对应 `--output-dir`、`--delay`，例如 `python scripts/demo/word.py --delay 0`。

| 应用 | 展示内容 |
| --- | --- |
| Word | 标题和正文排版、原生表格、内容读回、显式保存。 |
| Excel | 数据、公式和计算、格式、读回与保存。 |
| PPT | 幻灯片、文字和形状、复制页面、表格、演讲备注与保存。 |

演示使用唯一文件名和独立输出目录，不覆盖已有文件。空白文件由演示夹具准备，所有展示的编辑通过正式 Action Session 执行；这不代表新增了首次保存或另存为 Action。正常结束保留已保存的 WPS 文档窗口。

## 原生验证

以下入口需要实际 Windows WPS 环境：

```powershell
python scripts/validate/word.py --output-dir build/word-acceptance
python scripts/validate/excel.py --output-dir build/excel-acceptance --wps-version 12.0.0.28505
python scripts/validate/excel_common.py --output-dir build/excel-common-acceptance
python scripts/validate/ppt.py --output-dir build/ppt-acceptance
python scripts/validate/ppt_common.py --output-dir build/ppt-common-acceptance
```

Excel 的 `--wps-version` 应填写实际测试版本。Word 验证入口覆盖演示闭环、保存后的 DOCX 文字／表格以及精确文档窗口，不宣称覆盖全部 13 项 Word Action。Excel／PPT 分别保留基础和常用动作验收。

完整启动链的控制台观察程序位于 `src/test/python/tests/{word,excel,ppt}/desktop_acceptance.pyw`。不依赖 WPS 的回归统一运行：

```sh
PYTHONPATH=src/main/python python -m unittest discover -s src/test/python -p 'test_*.py'
```

原有平铺入口已迁入分类目录，命令路径以上表为准。
