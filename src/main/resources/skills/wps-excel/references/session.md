# Session Client

脚本路径相对于完整 Skill 目录。在 Windows 上将下面脚本中的目录和文件路径替换为用户授权的实际值。`call` 返回完整 Action Response，成功数据在 `data` 中；失败会抛出 `ActionFailed`，不能捕获后无条件继续写入。

```python
import sys
sys.path.insert(0, r"C:\skills\wps-excel\scripts")
from excel import open_session

with open_session() as client:
    def call(action, params):
        return client.call({"app": "excel", "action": action}, params)["data"]

    opened = call("openWorkbook", {"path": r"C:\work\report.xlsx"})
    info = call("getWorkbookInfo", {})
    page = call("listWorksheets", {"offset": 0, "limit": 100})
    # 根据用户目标和返回的工作表名称确定 sheet；此示例只读。
    region = call("readRange", {"sheet": "Sheet1", "address": "A1:B10"})
    print(info, page, region)
```

`listWorksheets` 返回 `total` 和 `nextOffset`；非 null 时继续翻页。工作表序号是观察事实，不是后续 Action 的目标地址。真实任务需要保存时，在内容验证通过后显式 `save`。退出上下文后可检查 `client.session_outcome` 和 `client.cleanup_error`。

`getWorkbookInfo.window` 返回已绑定工作簿的窗口句柄和所属进程，窗口不可用时为 null。此信息用于展示验证，不是选择工作簿的地址。WPS 顶层窗口可能统一显示“WPS Office”，不要依赖标题包含文件名来判断工作簿是否显示。

WPS 不允许同时打开文件名相同的两个工作簿，即使目录不同。`openWorkbook` 会先复用同一文件的既有绑定；若名称相同但文件身份不同，返回 `DOCUMENT_OPEN_FAILED`，避免阻塞对话框。改用不同文件名或由用户明确关闭冲突文件后再决定是否打开；不能把同名文件当成原目标。

如果 WPS 仅安装/注册到当前用户，管理员进程可能无法解析 `KET.Application`。使用普通 PowerShell；运行权限应与用户桌面 WPS 一致。桥接在准备阶段检查 COM 可用性，返回明确的 `EXCEL_CAPABILITY_UNAVAILABLE`，不通过硬编码 CLSID 或修改注册表绕过权限边界。
