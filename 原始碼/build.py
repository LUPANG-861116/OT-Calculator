#!/usr/bin/env python3
"""
把「介面.html」「邏輯.js」「jszip.min.js」三個檔合併成可以直接開的
單一檔「超勤格轉換器.html」。

用法：在「原始碼」資料夾裡執行
    python3 build.py

改完程式碼一定要重跑這支，不然成品檔還是舊的。
"""
import pathlib

HERE = pathlib.Path(__file__).parent
OUT = HERE.parent / "超勤格轉換器.html"

ui = (HERE / "介面.html").read_text(encoding="utf-8")
js = (HERE / "邏輯.js").read_text(encoding="utf-8")
zp = (HERE / "jszip.min.js").read_text(encoding="utf-8")
tpl_file = HERE / "templates.js"
tpl = tpl_file.read_text(encoding="utf-8") if tpl_file.exists() else ""

# 介面.html 的結構是「<title> + <link> + <style>」後面接「<div class="wrap">…」，
# 從第一個 <div class="wrap"> 切開，前半塞進 <head>，後半放進 <body>。
MARK = '\n<div class="wrap">'
head, rest = ui.split(MARK, 1)

content = head + MARK + rest + "\n<script>\n" + zp + "\n</script>\n<script>\n" + tpl + "\n</script>\n<script>\n" + js + "\n</script>\n"
head2, body2 = content.split(MARK, 1)

html = (
    '<!doctype html>\n<html lang="zh-Hant">\n<head>\n'
    '<meta charset="utf-8">\n'
    '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
    '<style>html{color-scheme:light dark}body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>\n'
    + head2 +
    '\n</head>\n<body>' + MARK + body2 + '\n</body>\n</html>\n'
)

OUT.write_text(html, encoding="utf-8")
print(f"已產生 {OUT}（{len(html):,} 字元）")
