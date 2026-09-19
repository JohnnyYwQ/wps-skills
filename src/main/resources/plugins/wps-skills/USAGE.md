# WPS Skills

让 Agent 在本机真实 Windows WPS 中创建、读取、编辑和交付办公文件。

| Skill | 能力 |
| --- | --- |
| [wps-word](skills/wps-word/SKILL.md) | 正文、排版、表格、图片、DOCX 保存与 PDF 导出。 |
| [wps-excel](skills/wps-excel/SKILL.md) | 工作表、数据、公式、区域格式、XLSX 保存与 PDF 导出。 |
| [wps-ppt](skills/wps-ppt/SKILL.md) | 幻灯片、文字、形状、表格、备注、PPTX 保存及 PDF/PNG 导出。 |

选择对应 Skill 并描述目标即可。Agent 按需查询操作定义，提交包含多个步骤的 Task JSON，读取回执后报告结果。每个 Task 绑定一份明确文档，结束后释放其执行资源。

需要 Windows、Python 3.10+、Windows PowerShell 5.1、相应 WPS 应用和已登录的桌面。三个 Skill 使用同一套任务机制，各自携带完整依赖，可独立工作。

输出使用本机绝对路径，父目录应已存在。仅在明确请求时保存或导出；失败或结果不确定时停止后续操作，保留已发生的效果，不自动回滚或重试。可按原请求路径查询持久回执。

本插件使用 [MIT License](LICENSE)。
