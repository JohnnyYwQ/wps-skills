# 新建、保存与导出

已有文件使用 `openDocument`；明确新建时使用 `createDocument`，不传路径。首次保存和另存为都是同一个 `saveAs` Action，参数为 Windows 主机上的绝对 `.docx` 路径和 `overwritePolicy: "failIfExists"`。

输出父目录必须已存在，目标文件必须不存在。同名文件、目录、当前文件及其硬链接不会被覆盖；另一个会话占用或隔离的目标也不能接管。已有授权仍不能使本版 `saveAs` 支持覆盖：换用用户接受的新路径，不能自行删除已有文件。

```python
saved = client.call({"app": "word", "action": "saveAs"}, {
    "outputPath": r"C:\work\new.docx",
    "overwritePolicy": "failIfExists",
})["data"]
assert saved["documentState"]["persistenceState"] == "saved"
```

成功后仍是同一个实际文档；可继续编辑，`save` 写入刚确认的新路径。旧文件保持另存为前的内容。旧路径及新旧文件身份的会话占用一直保留到清理，其他会话不能在迁移间隙插入。新建未保存时调用 `save` 会返回 `PERSISTENCE_LOCATOR_REQUIRED`。

PDF 使用独立的 `exportPdf`，保持 Content Revision 和保存状态；不会给新建文档建立 DOCX 路径。PDF 的覆盖策略以其单独契约为准，不能套用到 `saveAs`。

保存调用开始后的错误可能已经写出或改变了底层保存身份，返回 `unknown` 时不自动重试。会话终止后报告实际状态，不把打开新旧路径作为隐式恢复。清理不补做保存，也不关闭文档。
