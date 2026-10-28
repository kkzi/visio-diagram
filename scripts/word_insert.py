# -*- coding: utf-8 -*-
"""
word_insert —— 用 Word COM 把图片（含 EMF 矢量）插入 docx 的占位符位置。

这是推荐的「路线 A」：Word 原生支持 EMF，矢量保真，无需手拼 OOXML。

用法：
    # 把文档里所有 {{FIG:N2}} 换成图，宽度默认取正文栏宽
    python word_insert.py 方案.docx --placeholder "{{FIG:N2}}" --image 图N2.emf

    # 指定宽度（mm）与输出文件
    python word_insert.py in.docx --ph "{{FIG:N3}}" --image 图N3.emf -w 170 -o out.docx

也可作为库调用：
    from word_insert import insert_at_placeholder
    insert_at_placeholder("in.docx", "{{FIG:N2}}", "图N2.emf")
"""

import os
import sys
import time

import win32com.client as win32

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_embed import emf_size_mm, content_column_mm   # noqa: E402

MM2PT = 72.0 / 25.4
WD_ALIGN_CENTER = 1


def _retry(fn, times=6, delay=0.4):
    """Word 偶尔会抛 'Call was rejected by callee'（COM 忙），重试即可。"""
    for i in range(times):
        try:
            return fn()
        except Exception:
            if i == times - 1:
                raise
            time.sleep(delay)


def _quiet(fn):
    """忽略一切异常 —— 用于收尾（Close/Quit），失败不应掩盖主流程结果。"""
    try:
        fn()
    except Exception:
        pass


def insert_at_placeholder(docx, placeholder, image, width_mm=None,
                          out=None, center=True, replace_all=True):
    """把 docx 中所有 `placeholder` 文本替换为 `image`。返回插入张数。"""
    docx = os.path.abspath(docx)
    image = os.path.abspath(image)
    out = os.path.abspath(out or docx)

    if width_mm is None:
        col, _ = content_column_mm(docx)
        width_mm = (col - 5) if col else 150.0    # 默认略窄于栏宽，留呼吸感

    # 由 EMF 头取宽高比（PIL 读不了 EMF）
    sz = emf_size_mm(image) if image.lower().endswith(".emf") else None
    height_mm = (width_mm * sz[1] / sz[0]) if sz else None

    app = win32.DispatchEx("Word.Application")   # DispatchEx：独立新实例，
    #   避免接管用户已打开的 Word 实例（否则收尾 Quit() 会把用户的 Word 一起关掉）
    _quiet(lambda: setattr(app, "Visible", False))
    _quiet(lambda: setattr(app, "DisplayAlerts", 0))
    doc, n = None, 0
    try:
        doc = _retry(lambda: app.Documents.Open(docx, ReadOnly=False))
        while True:
            rng = doc.Content
            f = rng.Find
            f.ClearFormatting()
            f.Text = placeholder
            f.Forward = True
            f.Wrap = 0                      # wdFindStop
            if not f.Execute():
                break
            rng.Text = ""                   # 清掉占位符，rng 变为插入点
            if center:
                rng.ParagraphFormat.Alignment = WD_ALIGN_CENTER
            shp = _retry(lambda: doc.InlineShapes.AddPicture(
                FileName=image, LinkToFile=False, SaveWithDocument=True,
                Range=rng))
            shp.Width = width_mm * MM2PT
            if height_mm:
                shp.Height = height_mm * MM2PT
            n += 1
            if not replace_all:
                break
        if n:
            _quiet(lambda: doc.Fields.Update())
            _retry(lambda: doc.SaveAs(out))   # ★ 没插图就不重存：Word 保存会重排 XML（§26 类副作用）
    finally:
        if doc is not None:
            _quiet(lambda: doc.Close(False))
        _quiet(lambda: app.Quit())
    return n


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="把图片插入 docx 的占位符位置")
    ap.add_argument("docx")
    ap.add_argument("--placeholder", "--ph", dest="ph", required=True)
    ap.add_argument("--image", required=True)
    ap.add_argument("-w", "--width", type=float, default=None,
                    help="图片宽度 mm（默认 = 正文栏宽 − 5mm）")
    ap.add_argument("-o", "--out", default=None)
    a = ap.parse_args()
    cnt = insert_at_placeholder(a.docx, a.ph, a.image, a.width, a.out)
    print("已插入 %d 处：%s" % (cnt, a.out or a.docx))
