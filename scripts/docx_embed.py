# -*- coding: utf-8 -*-
"""
docx_embed —— 把 Visio 导出的 EMF 正确嵌进 Word 所需的底层工具。

包含三件事，都是踩过坑才总结出来的：
  1. emf_size_mm()        从 EMF 文件头读出图幅（PIL 读不了 EMF，必须手工解析）
  2. content_column_mm()  从 docx 的 sectPr 算出正文可用栏宽（**含装订线**）
  3. ensure_content_type()  往 [Content_Types].xml 里补 Default 扩展名声明

命令行自查：
    python docx_embed.py <docx>
    python docx_embed.py --emf <某个.emf>
"""

import os
import re
import shutil
import struct
import zipfile

EMU_PER_MM = 36000            # OOXML 长度单位：1 mm = 36000 EMU
TWIPS_PER_MM = 56.6929

EXT_MIME = {
    "emf": "image/x-emf",
    "wmf": "image/x-wmf",
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "svg": "image/svg+xml",
}


# ---------------------------------------------------------------- EMF
def emf_info(path):
    """解析 EMF 头，返回 dict。

    关键字段：
      bounds_px  rclBounds（偏移 8，4×int32）—— 设备单位（像素）
      frame_mm   rclFrame （偏移 24，4×int32，单位 0.01 mm）—— 物理尺寸
      size_mm    (宽mm, 高mm)，取自 rclFrame
      ratio      高/宽（**可靠**）
      implied_dpi bounds_px / frame_mm 推出的参考分辨率

    ⚠ 重要：EMF 的**绝对物理尺寸不可靠**。实测 Visio 导出的 EMF，
      rclFrame 隐含约 117 DPI，而同一张图的 PNG 是 96 DPI（物理尺寸差 ~22%）。
      只有**宽高比**在两种口径下完全一致（差异 < 0.1%）。
      因此嵌图时应**指定目标宽度**（按正文栏宽），高度按宽高比推算，
      不要直接采信 rclFrame 的绝对值。
    """
    with open(path, "rb") as f:
        head = f.read(56)
    if len(head) < 56 or head[:4] != b"\x01\x00\x00\x00":
        return None
    sig = head[40:44]
    if sig != b" EMF":
        return None
    bx = struct.unpack("<4i", head[8:24])
    fr = struct.unpack("<4i", head[24:40])
    w_px, h_px = bx[2] - bx[0], bx[3] - bx[1]
    w_mm, h_mm = (fr[2] - fr[0]) * 0.01, (fr[3] - fr[1]) * 0.01
    dpi = (w_px / (w_mm / 25.4)) if w_mm > 0 else 0
    return {
        "bounds_px": (w_px, h_px),
        "frame_mm": (w_mm, h_mm),
        "size_mm": (w_mm, h_mm),
        "ratio": (h_mm / w_mm) if w_mm > 0 else 0,
        "implied_dpi": dpi,
    }


def emf_size_mm(path):
    """读 EMF 图幅 (宽mm, 高mm)。取不到返回 None。

    注意这是 rclFrame 的绝对值，仅用于**取宽高比**；
    嵌图时请按目标宽度缩放（见 emf_info 的说明）。
    """
    info = emf_info(path)
    if not info:
        return None
    w, h = info["frame_mm"]
    if w <= 0 or h <= 0:
        return None
    return w, h


# ---------------------------------------------------------------- 版心宽度
def content_column_mm(docx_path, section=-1):
    """返回正文可用**单栏**宽度 (mm, twips)。

    = (页宽 − 左边距 − 右边距 − 装订线) / 栏数 − (栏数−1)/栏数 × 栏间距
    若带 w:gutterAtTop，则装订线在上边、不减宽度。

    ⚠ 装订线（w:gutter）在 Word 中是**生效**的，很多模板带 1.5 cm 装订线，
      按 21 cm 页宽算出的「17.5 cm 栏宽」就是这么来的；若忽略它会溢出页面。

    ⚠ 多栏：旧版只按单栏算，分栏文档会把「栏宽」算成整版宽。
      现读 `<w:cols w:num= w:space=>`，按栏数与栏间距折算**单栏**宽度。

    section：取第几个 `<w:sectPr>`。**默认 -1（最后一节）** —— 正文节通常在文末
      （封面/目录节在前），旧版默认读第 0 节会取错。要指定某节就传 0/1/...。
    """
    z = zipfile.ZipFile(docx_path)
    raw = z.read("word/document.xml").decode("utf-8")
    z.close()

    # 按 <w:sectPr> 切节（含自闭合写法）；一节都没有则整篇当一节
    secs = re.findall(r"<w:sectPr\b.*?</w:sectPr>|<w:sectPr\b[^>]*/>", raw, re.S)
    if not secs:
        secs = [raw]
    s = secs[section] if -len(secs) <= section < len(secs) else secs[-1]

    pg = re.search(r"<w:pgSz[^>]*/>", s)
    pm = re.search(r"<w:pgMar[^>]*/>", s)
    if not pg or not pm:
        return None, None
    pg, pm = pg.group(0), pm.group(0)

    def attr(src, name):
        m = re.search(r'w:%s="(-?\d+)"' % name, src)
        return int(m.group(1)) if m else 0

    w = attr(pg, "w") - attr(pm, "left") - attr(pm, "right")
    if "w:gutterAtTop" not in pm:
        w -= attr(pm, "gutter")

    # 分栏折算：num 缺省即 1 栏
    cols = re.search(r"<w:cols\b[^>]*/?>", s)
    num = attr(cols.group(0), "num") if cols else 0
    space = attr(cols.group(0), "space") if cols else 0
    if num and num > 1:
        w = int(round((w - space * (num - 1)) / float(num)))

    return w / TWIPS_PER_MM, w


# ---------------------------------------------------------------- Content_Types
def ensure_content_type(docx_path, ext, mime=None, out_path=None):
    """确保 [Content_Types].xml 里声明了该扩展名。返回 (是否改动, 输出路径)。

    ⚠ 坑：`<Default>` 必须写在根节点 `<Types>` **内部**。
      若插到 XML 声明之后、`<Types>` 之外，Word 会报「文件已损坏」。
    """
    ext = ext.lower().lstrip(".")
    mime = mime or EXT_MIME.get(ext)
    if not mime:
        raise ValueError("未知扩展名 %r，请显式传 mime" % ext)
    out_path = out_path or docx_path
    tmp = out_path + ".tmp"

    changed = False
    zin = zipfile.ZipFile(docx_path)
    names = zin.namelist()
    ct = zin.read("[Content_Types].xml").decode("utf-8")

    if re.search(r'<Default[^>]*Extension="%s"' % re.escape(ext), ct, re.I):
        changed = False
    else:
        m = re.search(r"<Types[^>]*>", ct)
        if not m:
            zin.close()
            raise RuntimeError("[Content_Types].xml 缺少 <Types> 根节点")
        decl = '<Default Extension="%s" ContentType="%s"/>' % (ext, mime)
        ct = ct[:m.end()] + decl + ct[m.end():]      # 插在 <Types> 内
        changed = True

    if not changed and os.path.abspath(out_path) == os.path.abspath(docx_path):
        zin.close()
        return False, docx_path

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            data = ct.encode("utf-8") if n == "[Content_Types].xml" else zin.read(n)
            zout.writestr(n, data)
    zin.close()
    shutil.move(tmp, out_path)
    return changed, out_path


# ---------------------------------------------------------------- 命令行自查
if __name__ == "__main__":
    import sys
    if len(sys.argv) >= 3 and sys.argv[1] == "--emf":
        p = sys.argv[2]
        info = emf_info(p)
        if not info:
            print("不是有效的 EMF：%s" % p)
        else:
            print("EMF       : %s" % os.path.basename(p))
            print("rclBounds : %d x %d px" % info["bounds_px"])
            print("rclFrame  : %.2f x %.2f mm" % info["frame_mm"])
            print("宽高比    : %.4f  （嵌图按此比例算高）" % info["ratio"])
            print("隐含 DPI  : %.1f  （⚠ 绝对值不可信，仅用比例）" % info["implied_dpi"])
    elif len(sys.argv) >= 2:
        p = sys.argv[1]
        mm, tw = content_column_mm(p)
        print("文档      : %s" % os.path.basename(p))
        if tw:
            print("正文栏宽  : %.2f cm  (%d twips)" % (mm / 10.0, tw))
        else:
            print("正文栏宽  : 未取到 sectPr")
    else:
        print(__doc__)
