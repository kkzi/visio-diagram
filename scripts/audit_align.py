# -*- coding: utf-8 -*-
"""
audit_align —— 箭头对齐审计（Visio vsdx）。

专治「箭头漂移」：箭头虽然连到了框，但落点不在框边中点附近，
看起来就是箭头歪了 / 快贴到框角了。

判据（只审带箭头的线段，天然排除图例线）：
  1. 取线段端点 P，找距离 P 最近、且 P 落在其跨度内的框边 E；
  2. 按 (框, 边) 归组；
  3. **只对「该边上只有一条箭头」的情况**要求居中 —— 宽框上多条箭头
     本来就分布在不同位置，属正常；偏移 > warn（默认 2mm）才报「漂移」。

★★ 框的识别按**形状类型**，不按「有没有文字」（2026-10-13 修正的重大缺陷）：
  旧版写 `if sh.Text: continue`（本意跳过纯文本框），但 `visio_lib.merge_texts()`
  **默认在保存前把标题并入矩形**，于是**几乎所有实体框都带字**，全被跳过 ——
  实测 图N5 0 个框、图N9 0 个框、图N3 只认出 2 个，却仍打印「漂移 0，exit 0」，
  审计形同失效。
  正确判据（本库自绘形状的稳定特征）：
    · 纯文本框（`VDraw.text()`）→ `LinePattern=0` 且 `FillPattern=0`（无边框无填充）；
    · 矩形 / 泳道框（`rect()/box()/step_box()/lane()`）→ 至少一个有值。
  再叠加**面积下限**（< 4mm² 的极小形状不当框）与「1-D 形状不参与」。
  ★ 另外**审计样本为 0 时不再静默"通过"**：若本图**有带箭头线段**却一条边都审不到
    → 判失败（exit 2，疑似框识别失败）；若本图**确实没有箭头**（纯布局图）→ 仅提示、
    不算失败。批量跑时按文件区分。

用法：
    python audit_align.py 图.vsdx [更多.vsdx...]
    python audit_align.py --warn 3 图.vsdx
    python audit_align.py --geom 图.vsdx      # 顺带打印全部框坐标

白名单（已定性为「刻意偏移」的端点，批量回归不再计失败）：
    python audit_align.py --add 图.vsdx       # 把本次报出的偏移端点登记进白名单
    白名单文件 = 脚本旁 audit_whitelist.json，按「图名（去扩展名）+ 端点坐标 + 边名」匹配；
    端点坐标允许 ±tol_mm 漂移（默认 3mm，防整体布局微移导致漏配；挪远了会重新报警，属预期）。
    已登记的命中打印为「白名单」；若某条目在图中**不再命中**，结尾提示「已失效，建议清理」。
"""

import argparse
import json
import os
import sys

# ★ Windows 控制台默认 GBK，打印 ★/✗ 等符号可能 UnicodeEncodeError → 退出码被改写。
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

MM = 25.4
MIN_BOX_MM2 = 4.0          # 面积下限：小于此不当「实体框」（滤掉装饰小方块）
WL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "audit_whitelist.json")
WL_TOL = 3.0               # 白名单端点匹配容差 mm（防布局微移漏配）


def _is_frame(sh):
    """判断 2-D 形状是否为「实体框」（矩形 / 泳道），而非纯文本框。

    ★ 不能用 `sh.Text` 判断（见模块 docstring）：`merge_texts()` 会把标题并入
      矩形，保存后几乎所有框都带字。
    本库自绘形状的稳定特征：
      · 纯文本框（`VDraw.text()`）→ LinePattern=0 **且** FillPattern=0；
      · 矩形 / 泳道 → 至少一个有值。
    保守起见：两个都取不到时**当作框**（宁可多审，不要漏审）。
    """
    lp = fp = None
    try:
        lp = sh.CellsU("LinePattern").ResultIU
    except Exception:
        pass
    try:
        fp = sh.CellsU("FillPattern").ResultIU
    except Exception:
        pass
    if lp is None and fp is None:
        return True
    return not ((lp or 0) == 0 and (fp or 0) == 0)


def _read(path):
    import win32com.client
    vis = win32com.client.DispatchEx("Visio.InvisibleApp")
    try:
        doc = vis.Documents.Open(os.path.abspath(path))
        pg = doc.Pages(1)
        PH = pg.PageSheet.CellsU("PageHeight").ResultIU * MM
        PW = pg.PageSheet.CellsU("PageWidth").ResultIU * MM
        boxes, arrows = [], []
        n_shape = n_text = n_skip_small = 0
        for sh in pg.Shapes:
            n_shape += 1
            try:
                if sh.OneD:
                    # 只收带箭头的线段；图例 / 装饰线 EndArrow=0 被排除
                    if sh.CellsU("EndArrow").ResultIU == 0:
                        continue
                    b = (sh.CellsU("BeginX").ResultIU * MM, PH - sh.CellsU("BeginY").ResultIU * MM)
                    e = (sh.CellsU("EndX").ResultIU * MM, PH - sh.CellsU("EndY").ResultIU * MM)
                    arrows.append((b, e))
                else:
                    if not _is_frame(sh):
                        n_text += 1                # 纯文本框：不算实体框
                        continue
                    x = sh.CellsU("PinX").ResultIU * MM
                    y = PH - sh.CellsU("PinY").ResultIU * MM
                    w = sh.CellsU("Width").ResultIU * MM
                    h = sh.CellsU("Height").ResultIU * MM
                    if w * h < MIN_BOX_MM2:
                        n_skip_small += 1
                        continue
                    boxes.append((x - w / 2, y - h / 2, x + w / 2, y + h / 2))
            except Exception:
                pass
        doc.Close()
    finally:
        vis.Quit()
    return boxes, arrows, PW, PH, {"shapes": n_shape, "text_skipped": n_text,
                                   "small_skipped": n_skip_small}


def _drop_lanes(boxes):
    """剔除泳道 / 分区大框（包含 >=2 个其他框者）。"""
    def contains(a, b):
        return (a[0] <= b[0] + .1 and a[1] <= b[1] + .1
                and a[2] >= b[2] - .1 and a[3] >= b[3] - .1)
    lanes = [a for a in boxes
             if sum(1 for b in boxes if b is not a and contains(a, b)) >= 2]
    return [b for b in boxes if b not in lanes]


def _nearest_edge(p, boxes, tol):
    """返回 (框索引, 垂直距离, 边名, 边中点, 跨度, 轴向) —— 取最近的一条边。"""
    best = None
    for i, (x1, y1, x2, y2) in enumerate(boxes):
        cand = [
            (abs(p[1] - y1), "上边", (x1 + x2) / 2.0, (x1, x2), "h"),
            (abs(p[1] - y2), "下边", (x1 + x2) / 2.0, (x1, x2), "h"),
            (abs(p[0] - x1), "左边", (y1 + y2) / 2.0, (y1, y2), "v"),
            (abs(p[0] - x2), "右边", (y1 + y2) / 2.0, (y1, y2), "v"),
        ]
        for d, side, mid, span, axis in cand:
            coord = p[0] if axis == "h" else p[1]
            if d <= tol and span[0] - tol <= coord <= span[1] + tol:
                if best is None or d < best[1]:
                    best = (i, d, side, mid, span, axis)
    return best


def _load_wl():
    if not os.path.isfile(WL_PATH):
        return {}
    try:
        with open(WL_PATH, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception as e:
        print("  !! 白名单文件读取失败（按空处理）：%s" % e)
        return {}


def _save_wl(wl):
    with open(WL_PATH, "w", encoding="utf-8") as fh:
        json.dump(wl, fh, ensure_ascii=False, indent=2)


def _wl_key(path):
    return os.path.splitext(os.path.basename(path))[0]


def audit(path, warn=2.0, tol=1.2, dump=False, wl=None):
    wl = wl if wl is not None else _load_wl()
    entries = list(wl.get(_wl_key(path), []))     # 拷贝；命中即弹出，剩下的=已失效
    boxes, arrows, _, _, st = _read(path)
    n_raw = len(boxes)
    boxes = _drop_lanes(boxes)

    # 归组：(框索引, 边名) -> [(标签, 点, 偏移)]
    groups = {}
    for b, e in arrows:
        for tag, p in (("起", b), ("末", e)):
            hit = _nearest_edge(p, boxes, tol)
            if hit is None:
                continue                 # 折线拐点 / 跨道中段，不算异常
            i, d, side, mid, span, axis = hit
            off = abs((p[0] if axis == "h" else p[1]) - mid)
            groups.setdefault((i, side), []).append((tag, p, off, mid, span))

    n_checked = sum(1 for v in groups.values() if len(v) == 1)
    print("=" * 70)
    print("%s   实体框 %d（原始 %d，滤泳道 %d） / 带箭头线段 %d"
          % (os.path.basename(path), len(boxes), n_raw,
             n_raw - len(boxes), len(arrows)))
    print("       形状总数 %d ｜ 纯文本框跳过 %d ｜ 极小形状跳过 %d ｜ 参与居中判定 %d 条边"
          % (st["shapes"], st["text_skipped"], st["small_skipped"], n_checked))
    print("=" * 70)
    if dump:
        for k, bb in enumerate(sorted(boxes, key=lambda r: (r[1], r[0]))):
            print("   框 [%6.1f,%6.1f]-[%6.1f,%6.1f] 中心(%6.1f,%6.1f)"
                  % (bb[0], bb[1], bb[2], bb[3], (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2))

    bad = n_white = 0
    for (i, side), items in sorted(groups.items(), key=lambda kv: (kv[0][0], kv[0][1])):
        if len(items) != 1:
            continue                     # 同一边多条箭头：分布是正常的，不判漂移
        tag, p, off, mid, span = items[0]
        # 相对判据：宽框上箭头不必强求绝对居中，偏移超过边长的 15% 才算漂移
        limit = max(warn, 0.15 * (span[1] - span[0]))
        if off <= limit:
            continue
        hit_wl = None
        for ent in entries:
            if (ent.get("side") == side
                    and abs(ent.get("x", 1e9) - p[0]) <= WL_TOL
                    and abs(ent.get("y", 1e9) - p[1]) <= WL_TOL):
                hit_wl = ent
                break
        if hit_wl is not None:
            entries.remove(hit_wl)
            n_white += 1
            print("  %s端(%6.1f,%6.1f) 单独连到[%s] 偏移 %5.1f mm  → 白名单（%s）"
                  % (tag, p[0], p[1], side, off, hit_wl.get("note", "已定性刻意")))
            continue
        print("  %s端(%6.1f,%6.1f) 单独连到[%s] 偏移 %5.1f mm"
              "（边中点 %.1f，边长 %.0f，阈值 %.1f）  ← 建议核对"
              % (tag, p[0], p[1], side, off, mid, span[1] - span[0], limit))
        bad += 1
    print("  待核对端点：%d 个 ｜ 白名单豁免 %d 个（绝对阈值 %.1f mm，且 > 边长 15%%）"
          % (bad, n_white, warn))
    if bad:
        print("  注：这是启发式筛查。若某框从同一侧「扇出」多条箭头（各自接到不同高度的框），")
        print("      或箭头刻意避开文字，则偏移属正常 —— 请结合渲染图目视判断。")
    # ★ 不再静默「通过」（旧版框识别失效仍打印 exit 0，是重大缺陷）：
    #   · groups 非空 → 箭头端点确实落在了识别到的框边上 → 框识别 OK（哪怕没有「单箭头接单边」）
    #   · groups 为空 **且** 本图有带箭头线段 → 没有任何端点贴到框边 → 极可能框识别失败 → 可疑
    #   · groups 为空 **且** 本图无箭头（纯布局图）→ 无需居中判定，仅提示，不算失败
    suspicious = False
    if not groups:
        if arrows:
            suspicious = True
            print("  ★★ 本图有 %d 条带箭头线段，却**没有任何端点落在识别到的框边上** ——"
                  " 极可能框识别失败，判失败。请人工确认。" % len(arrows))
        else:
            print("  ★ 本图没有带箭头的连接线（0 条），无需居中判定。")
    stale = []
    if entries:
        stale = entries
        print("  ☆ 白名单有 %d 条在本图未命中（布局已变或已修）—— 可编辑 %s 清理"
              % (len(entries), os.path.basename(WL_PATH)))
    return bad, n_checked, suspicious, stale


def _collect_new(path, warn, tol):
    """按 audit() 同一套判据收集「待登记」的偏移端点（只读 vsdx，不改白名单）。"""
    boxes, arrows, _, _, st = _read(path)
    boxes = _drop_lanes(boxes)
    groups = {}
    for b, e in arrows:
        for tag, p in (("起", b), ("末", e)):
            hit = _nearest_edge(p, boxes, tol)
            if hit is None:
                continue
            i, d, side, mid, span, axis = hit
            off = abs((p[0] if axis == "h" else p[1]) - mid)
            groups.setdefault((i, side), []).append((tag, p, off, mid, span))
    out = []
    for (i, side), items in groups.items():
        if len(items) != 1:
            continue
        tag, p, off, mid, span = items[0]
        limit = max(warn, 0.15 * (span[1] - span[0]))
        if off > limit:
            out.append({"x": round(p[0], 2), "y": round(p[1], 2),
                        "side": side, "off": round(off, 2),
                        "note": "待人工定性（--add 登记）"})
    return out


def main():
    ap = argparse.ArgumentParser(description="Visio 箭头对齐审计")
    ap.add_argument("vsdx", nargs="+")
    ap.add_argument("--warn", type=float, default=2.0, help="偏移报警阈值 mm")
    ap.add_argument("--tol", type=float, default=1.2, help="判定贴边的容差 mm")
    ap.add_argument("--geom", action="store_true", help="打印全部框坐标")
    ap.add_argument("--allow-empty", action="store_true",
                    help="允许「有箭头却审不到框边」而不报失败（默认报失败）")
    ap.add_argument("--add", action="store_true",
                    help="把本次报出的偏移端点登记进白名单（audit_whitelist.json）"
                         "并直接判通过；登记前请先人工确认属刻意设计")
    a = ap.parse_args()

    if a.add:
        wl = _load_wl()
        n_new = 0
        for p in a.vsdx:
            news = _collect_new(p, a.warn, a.tol)
            ent = wl.get(_wl_key(p), [])
            # ★ 先按 (边名, 坐标±WL_TOL) 去重：重复 --add 同一文件会在白名单堆重复条目，
            #   多余的每条之后都会被报「已失效」（自己造噪音）
            news = [n for n in news if not any(
                e.get("side") == n["side"]
                and abs(e.get("x", 1e9) - n["x"]) <= WL_TOL
                and abs(e.get("y", 1e9) - n["y"]) <= WL_TOL for e in ent)]
            if news:
                wl.setdefault(_wl_key(p), []).extend(news)
                n_new += len(news)
                print("[--add] %s：登记 %d 条" % (os.path.basename(p), len(news)))
            else:
                print("[--add] %s：无需登记（无偏移端点或已在白名单）" % os.path.basename(p))
        if n_new:
            _save_wl(wl)
            print("[--add] 白名单已写入 %s（共 %d 图 / %d 条）"
                  % (WL_PATH, len(wl), n_new))
        sys.exit(0)

    total = checked = 0
    suspicious = []
    stale_all = []
    for p in a.vsdx:
        b, c, s, stale = audit(p, warn=a.warn, tol=a.tol, dump=a.geom)
        total += b
        checked += c
        if s:
            suspicious.append(os.path.basename(p))
        for ent in stale:
            stale_all.append((_wl_key(p), ent))
    print("\n合计漂移端点 %d｜参与判定 %d 条边" % (total, checked))
    if suspicious and not a.allow_empty:
        print("★ 以下文件「有带箭头线段，却没有任何端点落在识别到的框边上」"
              "（疑似框识别失败）→ 判失败：")
        for f in suspicious:
            print("   -", f)
        print("  （确认确实如此时可加 --allow-empty 强制通过）")
        sys.exit(2)
    if total:
        print("  （把已定性为刻意的偏移登记进白名单可跑：--add 同一批文件）")
        sys.exit(1)
    if stale_all:
        print("☆ 白名单失效条目 %d 条（未影响本次判定，可手动清理 %s）"
              % (len(stale_all), os.path.basename(WL_PATH)))
        for key, ent in stale_all:
            print("   - %s (%.1f,%.1f %s)" % (key, ent["x"], ent["y"], ent["side"]))
    sys.exit(0)


if __name__ == "__main__":
    main()
