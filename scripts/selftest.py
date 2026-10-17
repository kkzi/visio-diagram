# -*- coding: utf-8 -*-
"""visio_lib 自测 —— 画一张覆盖关键图元的测试图，存盘后解析 vsdx 断言。

为什么需要它：本库最贵的几条结论（连接线 ObjType、几何点数、文字并入、溢出
自检）都是「改一处、坏另一处」的回归高发区。本脚本把它们固化成可一键复跑的
断言，改完库先跑它，再出正式图。

覆盖项（对应 references/pitfalls.md 的结论）：
  A 两端胶水**直线**连接线  → ObjType=2、几何 2 点            （§17 §21）
  B 两端胶水**折线**连接线  → ObjType=2、几何 = 折点数        （§17 §21）
  C **自由端**连接线        → 不加 ObjType（Visio 会夹取改坏） （§21）
  D 连接线**自带标签**      → ObjType=2 且 Text 落在线上       （§20 §23）
  E 文字并入宿主            → 形状数下降、宿主获得文字         （§18）
  F 文字溢出**自检**        → check() 能报出溢出               （§3）
  G 图例**不越界**          → legend() 按 w 均布、整行不超画布   （§22）
  H 标签**贴线**（显式）    → 文字中心落在线上 ~gap 处、可换侧  （§23）
  I 标签**贴线**（并入）    → 摆在线上方 5mm 的文本框被收进 ~2.3mm（§23）
  J 连线**出线方向**自检    → 目标在右却从左边出线，必须报出     （§24）
  K 合法**绕行**不误报      → 先向左绕再回头接右边，不报          （§24）

用法：
    python scripts/selftest.py            # 在临时目录里跑，跑完自动清理
    python scripts/selftest.py --keep     # 保留产物，便于用 Visio 打开排查

退出码：0 = 全部通过；1 = 有失败。**不依赖任何项目文件**，可独立复跑。
"""
import argparse
import os
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

# ★ Windows 控制台默认 GBK，末行打印 ✅/❌ 会抛 UnicodeEncodeError，
#   把「全过」的退出码改写成 1（自测假失败）。先切 UTF-8 输出。
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from visio_lib import VDraw, LBL_GAP          # noqa: E402

NS = "{http://schemas.microsoft.com/office/visio/2012/main}"


def parse_shapes(vsdx):
    """读 vsdx（本质是 zip）→ {shape_id: {objtype, geo, text, cells}}，只取顶层形状。"""
    with zipfile.ZipFile(vsdx) as z:
        xml = z.read("visio/pages/page1.xml").decode("utf-8")
    root = ET.fromstring(xml)
    holder = root.find(NS + "Shapes")
    out = {}
    for shp in holder.findall(NS + "Shape"):
        sid = int(shp.get("ID"))
        objtype = None
        cells = {}
        for c in shp.findall(NS + "Cell"):
            if c.get("N") == "ObjType":
                objtype = c.get("V")
            try:
                cells[c.get("N")] = float(c.get("V")) * 25.4      # 英寸 → mm
            except (TypeError, ValueError):
                pass
        geo = 0
        for sec in shp.findall(NS + "Section"):
            if sec.get("N") == "Geometry":
                geo = len(sec.findall(NS + "Row"))       # MoveTo + 各 LineTo
        t = shp.find(NS + "Text")
        txt = "".join(t.itertext()).strip() if t is not None else ""
        out[sid] = {"objtype": objtype, "geo": geo, "text": txt, "cells": cells}
    return out


def label_center(sh, H):
    """1-D 形状**自带文字**的中心（页坐标 mm，左上原点、Y 向下）。

    TxtPin 是形状**局部坐标**（原点 = Pin−LocPin，Y 向上），故要加回局部原点、
    再把 Y 翻回来。返回 None = 这个形状没写 TxtPin（没有自带文本）。
    """
    c = sh.get("cells", {})
    need = ("TxtPinX", "TxtPinY", "PinX", "LocPinX", "PinY", "LocPinY")
    if any(k not in c for k in need):
        return None
    ox = c["PinX"] - c["LocPinX"]
    oy = c["PinY"] - c["LocPinY"]
    return (c["TxtPinX"] + ox, H - (c["TxtPinY"] + oy))


def shape_extent(vsdx):
    """返回 (minx, maxx, miny, maxy)，单位 mm（vsdx 里是英寸，×25.4）。"""
    with zipfile.ZipFile(vsdx) as z:
        xml = z.read("visio/pages/page1.xml").decode("utf-8")
    root = ET.fromstring(xml)
    xs, ys = [], []
    for shp in root.find(NS + "Shapes").findall(NS + "Shape"):
        cell = {c.get("N"): c.get("V") for c in shp.findall(NS + "Cell")}
        try:
            px = float(cell.get("PinX", 0)) * 25.4
            py = float(cell.get("PinY", 0)) * 25.4
            w = float(cell.get("Width", 0)) * 25.4
            h = float(cell.get("Height", 0)) * 25.4
        except ValueError:
            continue
        xs += [px - w / 2, px + w / 2]
        ys += [py - h / 2, py + h / 2]
    return min(xs), max(xs), min(ys), max(ys)


def main():
    ap = argparse.ArgumentParser(description="visio_lib 自测")
    ap.add_argument("--keep", action="store_true", help="保留临时产物（便于排查）")
    args = ap.parse_args()

    tmp = tempfile.mkdtemp(prefix="visio_selftest_")
    vsdx = os.path.join(tmp, "selftest.vsdx")
    results = []
    instances = []          # 收集所有 VDraw，finally 里统一 close，防异常残留 VISIO.EXE

    def check(name, ok, detail=""):
        results.append(bool(ok))
        print("  [%s] %-26s %s" % ("PASS" if ok else "FAIL", name, detail))

    def safe_close():
        for d in instances:
            try:
                d.close()
            except Exception:
                pass

    try:
        d = VDraw(160, 90, theme="tech-cn")
        instances.append(d)

        # 两个方框，供连接线胶水
        a = d.rect(6, 10, 40, 22, fill="white")
        b = d.rect(100, 10, 40, 22, fill="white")
        cpa = {"r": d.add_cp(a, 1, 0.5), "b": d.add_cp(a, 0.5, 1),
               "t": d.add_cp(a, 0.5, 0)}
        cpb = {"l": d.add_cp(b, 0, 0.5), "b": d.add_cp(b, 0.5, 1),
               "t": d.add_cp(b, 0.5, 0), "r": d.add_cp(b, 1, 0.5)}

        # A 直线（两端胶水）—— 水平，顺带覆盖「Height 被 0.25in 夹取」路径
        idA = d.connect(a, cpa["r"], b, cpb["l"]).ID
        # B 折线（两端胶水，via 2 折点 → 4 点）
        idB = d.connect(a, cpa["b"], b, cpb["b"],
                        via=[(26, 32), (26, 62), (120, 62), (120, 32)]).ID
        # C 自由端（一端胶、一端固定点）→ 不应加 ObjType
        #   约定：不胶水那端把形状传 None、端点传 (x, y) 元组
        idC = d.connect(b, cpb["r"], None, (152, 74)).ID
        # D 连接线自带标签（两端胶水）
        idD = d.connect(a, cpa["t"], b, cpb["t"], label="回导数据").ID
        # D2 同路径再画一条，标签强制朝下 —— 覆盖 label_side
        idD2 = d.connect(a, cpa["t"], b, cpb["t"], label="比对",
                         label_side="down").ID
        # E 矩形 + 内部文本框（应被并入矩形）
        host = d.rect(60, 40, 40, 20, fill="white")
        idHost = host.ID
        d.text(64, 44, 32, 12, "宿主文字")

        n_drawn = d.pg.Shapes.Count
        d.save(vsdx)
        safe_close()          # 存完即关，后续只读 vsdx 解析，无需 Visio
        instances.clear()

        sh = parse_shapes(vsdx)
        g = lambda i: sh.get(i, {})

        check("A 直线 ObjType=2", g(idA).get("objtype") == "2",
              "objtype=%s" % g(idA).get("objtype"))
        check("A 直线几何 2 点", g(idA).get("geo") == 2,
              "geo=%s" % g(idA).get("geo"))
        check("B 折线 ObjType=2", g(idB).get("objtype") == "2",
              "objtype=%s" % g(idB).get("objtype"))
        check("B 折线几何 4 点", g(idB).get("geo") == 4,
              "geo=%s" % g(idB).get("geo"))
        check("C 自由端不加 ObjType", g(idC).get("objtype") != "2",
              "objtype=%s" % g(idC).get("objtype"))
        check("D 标签线 ObjType=2", g(idD).get("objtype") == "2",
              "objtype=%s" % g(idD).get("objtype"))
        check("D 标签落在连接线上", "回导数据" in g(idD).get("text", ""),
              "text=%r" % g(idD).get("text"))
        check("E 文字并入（形状数下降）", len(sh) < n_drawn,
              "绘制 %d → 存盘 %d" % (n_drawn, len(sh)))
        check("E 宿主获得文字", "宿主文字" in g(idHost).get("text", ""),
              "text=%r" % g(idHost).get("text"))

        # H 标签贴线（显式 label=）：文字中心必须落在「线上 ± label_gap」，
        #   而不是折线包围盒中心（肘形线那个点常常根本不在线上，§23）
        #   D/D2 的线 = (26,10)→(120,10) 水平段，默认 gap=2.2
        cD = label_center(g(idD), 90)
        check("H 默认标签在线上方 2.2mm",
              cD is not None and abs(cD[0] - 73) < 0.5 and abs(cD[1] - 7.8) < 0.5,
              "中心=%s 期望≈(73, 7.8)"
              % ("无 TxtPin" if cD is None else "%.2f,%.2f" % cD))
        cD2 = label_center(g(idD2), 90)
        check("H label_side=down 翻到下方",
              cD2 is not None and abs(cD2[0] - 73) < 0.5 and abs(cD2[1] - 12.2) < 0.5,
              "中心=%s 期望≈(73, 12.2)"
              % ("无 TxtPin" if cD2 is None else "%.2f,%.2f" % cD2))

        # F 溢出检测（独立实例，不存盘）
        d2 = VDraw(60, 30)
        instances.append(d2)
        d2.text(2, 2, 10, 4, "这是一段明显会溢出的长文字说明")
        bad = d2.check(verbose=False)
        safe_close()
        instances.clear()
        check("F 溢出能被检出", len(bad) >= 1, "检出 %d 处" % len(bad))

        # G 图例不越界（legend 须按 w 均布，而非按固定 46mm 文字框顺排）
        LW = 120.0
        d3 = VDraw(LW, 40)
        instances.append(d3)
        d3.legend(6, 24, LW - 12,
                  [("网络连接", None), ("授时信号", 4), ("回导数据", 2)])
        p3 = os.path.join(tmp, "legend.vsdx")
        d3.save(p3)
        safe_close()
        instances.clear()
        _, mx, _, _ = shape_extent(p3)
        check("G 图例不越画布", mx <= LW + 0.5, "最右 %.1f / 画布 %.0f" % (mx, LW))

        # I 并入路径「贴线归一」：独立文本框摆在线上方 5mm，并入连接线后应被
        #   收到「字形离線 0.8mm」= LBL_GAP + 半字高（§23）
        IW = 140.0
        d4 = VDraw(IW, 40)
        instances.append(d4)
        ra = d4.rect(10, 15, 20, 10, fill="white")
        rb = d4.rect(110, 15, 20, 10, fill="white")
        cpa4 = {"r": d4.add_cp(ra, 1, 0.5)}
        cpb4 = {"l": d4.add_cp(rb, 0, 0.5)}
        d4.connect(ra, cpa4["r"], rb, cpb4["l"], arrow=0)     # 线 y=20，x 30..110
        d4.text(56, 12.5, 8, 5, "标签", size=7)               # 中心 (60,15)，离线 5mm
        p4 = os.path.join(tmp, "snap.vsdx")
        d4.save(p4)
        safe_close()
        instances.clear()
        sh4 = parse_shapes(p4)
        tgt = None
        for s4 in sh4.values():
            if "标签" in s4.get("text", ""):
                tgt = label_center(s4, 40)
        exp_y = 20.0 - (LBL_GAP + VDraw._lh(7.0) / 2.0)
        d_off = None if tgt is None else 20.0 - tgt[1]
        check("I 并入标签被贴到线旁 %.1fmm" % (20.0 - exp_y),
              tgt is not None and abs(tgt[0] - 60) < 0.6 and abs(tgt[1] - exp_y) < 0.6,
              "中心=%s 期望≈(60, %.2f)"
              % ("未并入" if tgt is None else "%.2f,%.2f" % tgt, exp_y))

        # J 连线方向自检（§24）：目标在**右**，却从 a 的**左边**出线
        #   → connect() 当场记一条 warn；audit_links() 还应报「穿入 a 内部」
        d5 = VDraw(160, 60)
        instances.append(d5)
        a5 = d5.rect(6, 10, 40, 22, fill="white")
        b5 = d5.rect(100, 10, 40, 22, fill="white")
        ca5 = d5.side_cps(a5, "left right")
        cb5 = d5.side_cps(b5, "left")
        d5.connect(a5, ca5["left"], b5, cb5["left"])
        w5 = list(d5._link_warns)
        cross5 = d5.audit_links(verbose=False)
        safe_close()
        instances.clear()
        check("J 错误出线被检出", len(w5) == 1 and "left" in w5[0],
              "warn=%d %s" % (len(w5), (w5[0][:24] + "…") if w5 else ""))
        check("J 同时报穿越实体框", len(cross5) >= 1, "cross=%d" % len(cross5))

        # K 合法绕行：从左边出线、先向左绕再回头 → **不应**误报（§24）
        d6 = VDraw(160, 60)
        instances.append(d6)
        a6 = d6.rect(6, 20, 40, 22, fill="white")     # y 20~42
        b6 = d6.rect(100, 10, 40, 22, fill="white")   # y 10~32
        ca6 = d6.side_cps(a6, "left")
        cb6 = d6.side_cps(b6, "left")
        d6.connect(a6, ca6["left"], b6, cb6["left"],
                   via=[(2, 31), (2, 16), (100, 16), (100, 21)])
        w6 = list(d6._link_warns)
        cross6 = d6.audit_links(verbose=False)
        safe_close()
        instances.clear()
        check("K 合法绕行不误报", not w6 and not cross6,
              "warn=%d cross=%d" % (len(w6), len(cross6)))

    except Exception:
        import traceback
        traceback.print_exc()
        results.append(False)
    finally:
        safe_close()          # 兜底：异常退出也关
        if args.keep:
            print("\n  产物保留在：%s" % tmp)
        else:
            shutil.rmtree(tmp, ignore_errors=True)

    n_ok = sum(results)
    ok = bool(results) and n_ok == len(results)
    print("\n%s  通过 %d/%d" % ("✅ 全部通过" if ok else "❌ 有失败",
                                n_ok, len(results)))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
