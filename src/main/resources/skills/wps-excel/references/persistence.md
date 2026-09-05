# 新建、保存与导出

已有文件使用 `openWorkbook`；明确新建时使用 `createWorkbook`，不传路径。首次保存和另存为都是同一个 `saveAs` Action，参数为 Windows 主机上的绝对 `.xlsx` 路径和 `overwritePolicy: "failIfExists"`。

输出父目录必须已存在，目标文件必须不存在。同名文件、目录、当前文件及其硬链接不会被覆盖；另一个会话占用或隔离的目标也不能接管。已有授权仍不能使本版 `saveAs` 支持覆盖：换用用户接受的新路径，不能自行删除已有文件。

```python
saved = client.call({"app": "excel", "action": "saveAs"}, {
    "outputPath": r"C:\work\new.xlsx",
    "overwritePolicy": "failIfExists",
})["data"]
assert saved["documentState"]["persistenceState"] == "saved"
```

成功后仍是同一个实际文档；可继续编辑，`save` 写入刚确认的新路径。旧文件保持另存为前的内容。旧路径及新旧文件身份的会话占用一直保留到清理，其他会话不能在迁移间隙插入。新建未保存时调用 `save` 会返回 `PERSISTENCE_LOCATOR_REQUIRED`。

PDF 导出使用 `exportPdf`，传入 `.pdf` 的 `outputPath` 和 `overwritePolicy: "failIfExists"`。它按 WPS 当前原生导出／打印设置输出，验证 PDF 文件头及导出前后的内容观察、保存状态和绑定路径；不保存源文件。检查结果中的 `documentStateBefore` 与 `documentStateAfter`，并按交付要求检查 PDF 页面。新建但未保存的文档也可以导出。

`saveAs` 和 `exportPdf` 当前验证范围最多 20 张工作表，每张 UsedRange 最多 1000 个单元格；超过范围会在写出前拒绝。单元格读回覆盖值、公式和已支持的格式，不代表图表或未读取的工作簿特性已完整验证。

保存调用开始后的错误可能已经写出或改变了底层保存身份，返回 `unknown` 时不自动重试。会话终止后报告实际状态，不把打开新旧路径作为隐式恢复。清理不补做保存，也不关闭文档。
