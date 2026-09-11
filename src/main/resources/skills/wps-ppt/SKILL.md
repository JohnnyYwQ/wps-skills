---
name: wps-ppt
description: 通过 Windows WPS 演示新建或编辑 PPTX，管理幻灯片、图片和表格，调整文字排版、背景与备注，保存、另存为、导出 PDF 和单页 PNG。用户要求在可见 WPS 中操作演示文稿时使用。
---

# WPS PPT

使用本 Skill 的 `scripts/ppt.py` 和 PPT Application Contract Set。路径相对于本文件；先定位 Skill 绝对目录。执行需要 Windows、Python 3.8+、Windows PowerShell 与注册为 `KWPP.Application` 的 WPS 演示，不需要 pywin32。保留完整 `runtime/`；能力查询可在其他系统执行。

## 文稿与保存意图

已有文稿使用 Windows 主机上 `.pptx` 的绝对路径调用 `openPresentation`；明确新建时使用 `createPresentation`，得到一个没有幻灯片的未保存文稿，再用 `listSlides` 和 `addSlide` 添加页面。一个 Session 只绑定一个文稿，不选择活动窗口，不把打开失败改为新建。当前不提供图表、动画或旧版 `.ppt` / `.dps` / 宏文件支持。

修改已有文件默认需要显式 `save`，除非用户要求保留未保存状态。若 `openPresentation` 返回 `modified` 且本次需要保存，在首次修改前说明原有未保存修改也会一起保存，并取得确认；沿用用户已有的明确授权。只读任务不保存。

首次保存或另存为前读 [references/persistence.md](references/persistence.md)。`saveAs` 只写入尚不存在的 `.pptx`，之后 `save` 保存当前路径。导出不会替代源文件保存；用户未提供新建文件的输出路径时，可按意图保留未保存状态并如实报告。

## 发现与执行

```powershell
python "<skill-dir>/scripts/ppt.py" --app ppt --index
python "<skill-dir>/scripts/ppt.py" --app ppt --resolve openPresentation listSlides getSlideInfo setShapeText save
```

查询不启动 WPS。以解析的完整契约为准；`WPS_DISCOVERY_UNAVAILABLE` 表示尚未生产准入，不调用候选契约，也不绕过 Runtime 直接操作 COM。

执行前读 [references/session.md](references/session.md)。直接调用包内 `scripts/ppt.py` 的 `--start`、`--call`、`--status`、`--close`；Agent 只提供 Action 和 JSON 参数，无需编写任务脚本。一个任务复用返回的 `handle`，每次读取完整响应后才使用 `nextStep` 决定下一步。

先 `getPresentationInfo`、`listSlides`，再按返回的 `slideId` 调用 `getSlideInfo`。幻灯片位置是 1 起始的顺序；内容定位使用原生 ID，避免移动后按旧位置改错页。形状 ID 只在指定幻灯片内解释。

编辑前读 [references/editing.md](references/editing.md)，从指定读取动作取得 `token`，传入 `expectedToken`。检查返回的实际内容与用户目标。形状样式、对齐、查找替换、图片、表格、背景和备注操作读 [references/common.md](references/common.md)，使用各自对应的观察 token。完成后显式保存并验证 artifact、保存状态和最终文稿内容。

## 异常与结束

命令 JSON 中的 `response.error` 保留失败事实。`STALE_CONTENT` 需要重新读取并重新判断修改；`unknown` 或响应丢失表示可能已经部分执行，不自动重试。文稿关闭、绑定身份变化或桥接损坏会结束会话，不跟随新路径、不绕过 Lease 或 Quarantine。

结束 Session 只释放资源，不保存或关闭文稿。报告实际修改、验证、保存路径以及窗口仍然打开的状态；分别判断文稿结果与会话清理结果。

运行日志位置和自定义方法见 [references/logging.md](references/logging.md)。
