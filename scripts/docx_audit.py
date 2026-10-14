# -*- coding: utf-8 -*-
"""
docx_audit —— docx 交付前自查。静态检查（不需要 Word）+ 可选 Word 版式实测。

覆盖本次实战中踩到的全部坑：
  ✗ [Content_Types].xml 的 <Default> 写到了 <Types> 外面     → Word 报「文件已损坏」
  ✗ 根元素 mc:Ignorable 引用了未声明的 xmlns 前缀            → Word 报「文件已损坏」
  ✗ w:tblW / w:tcW 误用属性 w:val（正确是 w:w）              → 列宽塌缩、页数暴涨
  ✗ 表格宽度超过正文栏宽（常因忽略装订线 gutter）             → 右侧溢出页面
  ✗ 正文/表格残留加粗

用法：
    python docx_audit.py 方案.docx              # 静态 + Word 实测
    python docx_audit.py 方案.docx --no-word    # 只做静态检查
"""

import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

# ★ Windows 控制台默认 GBK，中文输出虽可过，但统一切 UTF-8 更稳（防符号类字符炸编码）。
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_embed import content_column_mm   # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


# ---------------------------------------------------------------- 静态检查
def static_check(path):
    """返回 (问题列表, 信息字典)。问题列表为空表示通过。"""
    problems, info = [], {}
    z = zipfile.ZipFile(path)
    names = z.namelist()

    # 1) 所有 XML 部件良构
    bad_xml = []
    for n in names:
        if n.endswith((".xml", ".rels")):
            try:
                ET.fromstring(z.read(n))
            except Exception as e:
                bad_xml.append("%s: %s" % (n, e))
    if bad_xml:
        problems.append("XML 不良构：\n    " + "\n    ".join(bad_xml))
    info["xml_parts"] = sum(1 for n in names if n.endswith((".xml", ".rels")))

    # 2) [Content_Types].xml：<Default> 必须在 <Types> 内
    ct = z.read("[Content_Types].xml").decode("utf-8", "ignore")
    i_types = ct.find("<Types")
    j_types = ct.find(">", i_types) if i_types >= 0 else -1
    if i_types < 0:
        problems.append("[Content_Types].xml 缺少 <Types> 根节点")
    else:
        before = ct[:i_types]
        if "<Default" in before or "<Override" in before:
            problems.append("[Content_Types].xml 的 <Default>/<Override> 出现在 <Types> 之外")
        # 图片扩展名是否都有声明
        exts = {os.path.splitext(n)[1].lower().lstrip(".")
                for n in names if n.startswith("word/media/")}
        declared = {m.lower() for m in re.findall(r'Extension="([^"]+)"', ct)}
        miss = sorted(exts - declared)
        if miss:
            problems.append("word/media 中的扩展名未在 Content_Types 声明：%s" % miss)
        info["media_exts"] = sorted(exts)

    # 3) mc:Ignorable 前缀必须已声明（否则 Word 报文件损坏）
    for n in ("word/document.xml", "word/styles.xml"):
        if n not in names:
            continue
        s = z.read(n).decode("utf-8", "ignore")
        m = re.search(r"<([\w:]+)[^>]*>", s)
        if not m:
            continue
        head = s[:m.end()]
        declared = set(re.findall(r'xmlns:([\w\d]+)="', head))
        ig = re.search(r'mc:Ignorable="([^"]+)"', head)
        if ig:
            und = [p for p in ig.group(1).split() if p not in declared]
            if und:
                problems.append("%s 的 mc:Ignorable 引用了未声明前缀：%s" % (n, und))

    # 4) tblW / tcW 属性名（必须是 w:w，不是 w:val）
    doc = z.read("word/document.xml").decode("utf-8", "ignore")
    for tag in ("tblW", "tcW"):
        bad = re.findall(r"<w:%s[^>]*?w:val=" % tag, doc)
        if bad:
            problems.append("w:%s 误用 w:val（应为 w:w），共 %d 处 → 列宽会塌缩、页数暴涨"
                            % (tag, len(bad)))

    # 5) 加粗 run 计数
    bolds = [b for b in re.findall(r"<w:b(?:\s[^>]*)?/>", doc)
             if 'w:val="0"' not in b and 'w:val="false"' not in b]
    info["bold_runs"] = len(bolds)
    if bolds:
        problems.append("document.xml 中存在 %d 个加粗 run" % len(bolds))

    # 6) 图片数量
    media = [n for n in names if n.startswith("word/media/")]
    info["media"] = [(os.path.basename(n), z.getinfo(n).file_size) for n in media]
    info["drawings"] = len(re.findall(r"<w:drawing>", doc))

    z.close()

    # 7) 正文栏宽
    mm, tw = content_column_mm(path)
    info["column_mm"] = mm
    if mm:
        info["column_cm"] = round(mm / 10.0, 2)
    return problems, info


# ---------------------------------------------------------------- Word 实测
def word_check(path):
    """用 Word COM 实测页数 / 表格宽度 / 图片尺寸 / 粗体。返回 (问题列表, 信息)。"""
    import win32com.client as win32
    problems, info = [], {}
    app = win32.DispatchEx("Word.Application")   # DispatchEx：独立新实例，
    #   不接管用户已打开的 Word（否则 Quit() 会连带关掉用户文档）
    app.Visible = False
    app.DisplayAlerts = 0
    doc = app.Documents.Open(os.path.abspath(path), ReadOnly=True)
    try:
        doc.Repaginate()
        info["pages"] = doc.ComputeStatistics(2)
        info["paragraphs"] = doc.ComputeStatistics(4)
        info["tables"] = doc.Tables.Count

        maxw, over = 0.0, []
        for ti in range(1, doc.Tables.Count + 1):
            t = doc.Tables(ti)
            tw = 0.0
            for ci in range(1, t.Columns.Count + 1):
                try:
                    tw += t.Columns(ci).Width
                except Exception:
                    pass
            cm = tw / 28.3465
            maxw = max(maxw, cm)
            if cm > (info.get("column_cm", 17.5) + 0.05):
                over.append((ti, round(cm, 2)))
        info["max_table_cm"] = round(maxw, 2)
        if over:
            problems.append("表格超宽（> 正文栏宽）：%s" % over[:8])

        info["images"] = []
        for i in range(1, doc.InlineShapes.Count + 1):
            s = doc.InlineShapes(i)
            info["images"].append((round(s.Width / 28.3465, 1),
                                   round(s.Height / 28.3465, 1)))

        bt = bo = 0
        for p in doc.Paragraphs:
            rng = p.Range
            if not rng.Bold:
                continue
            try:
                if rng.Information(12):      # wdWithInTable
                    bt += 1
                else:
                    bo += 1
            except Exception:
                bo += 1
        info["bold_in_table"] = bt
        info["bold_outside"] = bo            # 正常应等于「标题段落数」
        if bt:
            problems.append("表格内存在 %d 个粗体段落" % bt)
    finally:
        doc.Close(False)
        app.Quit()
    return problems, info


# ---------------------------------------------------------------- 主程序
def main():
    import argparse
    ap = argparse.ArgumentParser(description="docx 交付前自查")
    ap.add_argument("docx")
    ap.add_argument("--no-word", action="store_true", help="跳过 Word 实测")
    a = ap.parse_args()

    print("=" * 62)
    print("静态检查：%s" % os.path.basename(a.docx))
    print("=" * 62)
    problems, info = static_check(a.docx)
    print("XML 部件 %d · 图片 %d 张 · 加粗 run %d"
          % (info.get("xml_parts", 0), len(info.get("media", [])),
             info.get("bold_runs", 0)))
    if info.get("column_cm"):
        print("正文栏宽 %.2f cm（已计入装订线）" % info["column_cm"])
    for n, sz in info.get("media", []):
        print("   media: %s  %.0f KB" % (n, sz / 1024.0))

    if not a.no_word:
        print("-" * 62)
        print("Word 实测")
        print("-" * 62)
        try:
            p2, i2 = word_check(a.docx)
            problems += p2
            print("页数 %d · 段落 %d · 表格 %d"
                  % (i2.get("pages", 0), i2.get("paragraphs", 0), i2.get("tables", 0)))
            print("最宽表格 %.2f cm（栏宽 %.2f cm）"
                  % (i2.get("max_table_cm", 0), i2.get("column_cm", 0)))
            print("内嵌图片 %d 张" % len(i2.get("images", [])))
            for w_, h_ in i2.get("images", []):
                print("   %.1f x %.1f cm" % (w_, h_))
            print("粗体：表格内 %d · 表格外 %d（表格外正常=标题数）"
                  % (i2.get("bold_in_table", 0), i2.get("bold_outside", 0)))
        except Exception as e:
            print("Word 实测跳过：%s" % e)

    print("=" * 62)
    if problems:
        print("发现 %d 个问题：" % len(problems))
        for i, p in enumerate(problems, 1):
            print("  %d) %s" % (i, p))
        sys.exit(1)
    print("全部检查通过。")


if __name__ == "__main__":
    main()
