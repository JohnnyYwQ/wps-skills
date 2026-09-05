# 一个文稿、一个会话

在 Windows 普通桌面终端运行。管理员进程可能看不到当前用户的 COM 注册；不要因此修改注册表或重装 WPS。可以通过 SSH 在已授权的交互桌面测试，但 SSH 会话 0 本身不能证明窗口可见。

下面示例只读检查用户已指定的文件。把 Skill scripts 目录加入路径后导入薄入口：

```python
import sys
sys.path.insert(0, r"<skill-dir>\scripts")
from ppt import open_session, ActionFailed

with open_session() as client:
    opened = client.call({"app": "ppt", "action": "openPresentation"},
                         {"path": r"C:\work\slides.pptx"})["data"]
    slides = client.call({"app": "ppt", "action": "listSlides"}, {})["data"]
    if slides["slides"]:
        first = client.call({"app": "ppt", "action": "getSlideInfo"},
                            {"slideId": slides["slides"][0]["id"]})["data"]
        print(first)
```

修改时先根据用户授权及已有 dirty state 决定是否继续，再传观察 token。每个调用只有一个动作；不预先管道提交动作序列。`save` 只保存当前绑定的已有文件，无输出路径参数。清理后 WPS 文稿保持打开。

`getPresentationInfo.window` 来自所绑定文稿的原生窗口接口；null 表示没有可验证的窗口句柄。即使有句柄，也需要检查其所属进程和真实可见窗口尺寸，不能只匹配窗口标题。
