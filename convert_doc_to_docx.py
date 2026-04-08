"""批量把 .doc 转成 .docx（调用本机 Word，递归子目录）。

用法:
  python convert_doc_to_docx.py <目录>
  或把文件夹拖到本脚本图标上

已存在同名 .docx 时跳过；原 .doc 不会被删除。
"""
import sys
from pathlib import Path

# --- 安全开关：必须显式传目录，双击不执行 ---
if len(sys.argv) < 2:
    print("用法: 把要转换的文件夹拖到此脚本上，或在命令行传入目录路径。")
    input("按回车退出...")
    sys.exit(0)

ROOT = Path(sys.argv[1])
if not ROOT.is_dir():
    print(f"不是有效目录: {ROOT}")
    input("按回车退出...")
    sys.exit(1)

# --- 预览 + 二次确认 ---
docs = [p for p in ROOT.rglob("*.doc") if p.suffix.lower() == ".doc"]
if not docs:
    print(f"{ROOT} 下没有找到 .doc 文件")
    input("按回车退出...")
    sys.exit(0)

print(f"将在 {ROOT} 下递归转换 {len(docs)} 个 .doc 文件：")
for p in docs[:10]:
    print(" -", p.relative_to(ROOT))
if len(docs) > 10:
    print(f" ... 共 {len(docs)} 个")

if input("确认继续？(y/N) ").strip().lower() != "y":
    print("已取消")
    input("按回车退出...")
    sys.exit(0)

# --- 执行转换 ---
import win32com.client
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = False

ok = skip = fail = 0
try:
    for doc in docs:
        out = doc.with_suffix(".docx")
        if out.exists():
            print("skip:", doc.name)
            skip += 1
            continue
        try:
            d = word.Documents.Open(str(doc.resolve()), ReadOnly=True)
            d.SaveAs2(str(out.resolve()), FileFormat=16)  # wdFormatDocumentDefault
            d.Close(False)
            print("ok  :", doc.name)
            ok += 1
        except Exception as e:
            print("fail:", doc.name, e)
            fail += 1
finally:
    word.Quit()

print(f"\n完成：成功 {ok}，跳过 {skip}，失败 {fail}")
input("按回车退出...")
