# doc2docx

批量把 Word 97-2003 的 `.doc` 文件转换成 `.docx`，调用本机 Microsoft Word 完成转换，递归处理子目录。

## 为什么需要它

很多文档解析 / 向量化工具（`python-docx`、LangChain `Docx2txtLoader` 等）只支持 `.docx`，遇到老的 `.doc` 二进制格式会直接失败。这个脚本把目录下所有 `.doc` 一次性转成 `.docx`，原文件保留不动。

> 注意：这不是"编码转换"。`.doc` 是二进制格式，问题出在格式不兼容，而不是 UTF-8。

## 依赖

- Windows + 已安装 Microsoft Word
- Python 3
- `pywin32`

```bash
pip install pywin32
```

## 用法

**方式一：拖拽**

把要转换的文件夹拖到 `convert_doc_to_docx.py` 图标上。

**方式二：命令行**

```bash
python convert_doc_to_docx.py "D:\some\folder"
```

脚本会：

1. 递归扫描该目录下所有 `.doc`
2. 列出待转换文件清单，等待你输入 `y` 确认
3. 调用 Word 在原位置生成同名 `.docx`
4. 已存在的 `.docx` 自动跳过；原 `.doc` 不删除
5. 结束打印 成功 / 跳过 / 失败 统计

**安全设计**：直接双击脚本不会执行任何转换，只会提示用法后退出，避免误操作整盘文件。

## License

MIT
