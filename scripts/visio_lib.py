# -*- coding: utf-8 -*-
"""
visio_lib —— 通过 COM 驱动本机 Visio 绘制矢量图。

产物三件套：
    *.vsdx  原生可编辑（甲方/同事可继续在 Visio 里改）
    *.emf   矢量，供 Word / PPT 嵌入，缩放不失真
    *.png   预览，用于沟通与自查

坐标系约定（重要）
------------------
所有公开方法的坐标都以 **页面左上角为原点、Y 轴向下、单位 mm**。
库内部自动换算成 Visio 的「左下角原点 / Y 轴向上」。

用法
----
    from visio_lib import VDraw
    d = VDraw(150, 120, theme="tech-cn")
    d.box(5, 5, 60, 20, "标题", "说明文字")
    d.export_all("out/图A")
    d.check()          # 文字溢出自检
    d.close()

依赖
----
    pip install pywin32 pillow
    本机需安装 Visio（桌面版）。
"""

import math
import os
import re
import win32com.client as win32

import theme as _theme

MARGIN_MM = 0.8          # 图元内文字左右内边距
LBL_GAP = 0.8            # 连接线标签的**字形**到线段的距离（mm）
_FONTS = {}


def check_env(verbose=True):
    """开工前置检查：把「跑不起来」的问题在**最前面**就报出来。

    为什么要单独做（2026-10-13 实测踩坑）：
      · `import PIL` 通过 ≠ 能用 —— `PIL/__init__.py` 不加载 `ImageFont` 的 C 扩展，
        在 **Python 3.15 alpha** 下 `import PIL` 正常、`from PIL import ImageFont`
        直接崩（`slot ID 85`），而崩点会推迟到第一次画图量字宽时才暴露；
      · Visio 未安装 / 3.15 预览版 / Pillow 缺失，都应在开工前一次讲清。

    返回 (ok: bool, problems: list[str])。**只检查、不抛异常**，便于调用方自定策略。
    """
    import sys as _sys
    problems = []
    if _sys.version_info[:2] >= (3, 14):
        problems.append(
            "Python %d.%d 为预览/新版，Pillow 的 ImageFont 可能直接崩溃"
            "（实测 3.15 alpha + Pillow 12.3 → slot ID 85）。建议改用 3.9~3.13。"
            % _sys.version_info[:2])
    try:
        import win32com.client                    # noqa: F401
    except Exception as e:
        problems.append("缺少 pywin32（COM 依赖）：%s" % e)
    try:
        from PIL import ImageFont                 # noqa: F401
    except Exception as e:
        problems.append(
            "Pillow 的 ImageFont 不可用：%s\n"
            "    （注意：单独的 `import PIL` 能过，但它不加载 ImageFont 扩展，"
            "必须用 `from PIL import ImageFont` 检查）" % e)
    try:
        from PIL import __version__ as _pv
        if verbose:
            print("  环境：Python %s · Pillow %s"
                  % (_sys.version.split()[0], _pv))
    except Exception:
        pass
    if verbose:
        for p in problems:
            print("  !! " + p)
        if not problems:
            print("  前置检查通过。")
    return (not problems), problems


def _font(size_pt, font_name="微软雅黑"):
    key = (round(size_pt, 2), font_name)
    if key not in _FONTS:
        try:
            from PIL import ImageFont
        except Exception as e:                    # pragma: no cover
            raise RuntimeError(
                "Pillow 的 ImageFont 不可用（%s）。\n"
                "★ 先用 `python -c \"from PIL import ImageFont\"` 复现；"
                "若在 Python 3.14+ 预览版上失败，请改用 3.9~3.13。" % e)
        path = _theme.font_file(font_name, warn=True)
        if not path:
            raise RuntimeError("找不到可用字体文件，请在 theme.FONT_FILES 中补充")
        px = max(6, int(round(size_pt * 96.0 / 72.0)))
        _FONTS[key] = ImageFont.truetype(path, px)
    return _FONTS[key]


def _font_scale(size_pt, font_name="微软雅黑"):
    """返回 (font, scale)。★ PIL 的 truetype 会把字号**取整**（11.33→11px），
    使量出的字宽系统性偏小 ~3% → 折行行数少算（实测 图N6「读时刻（时间戳计数寄存器）」
    13 字在 38.4mm 宽里 Visio 折 2 行、PIL 却算 1 行）→ 合并时块高算小、文字下移 1.8mm。
    故按「真实 px / 取整 px」把长度放大回真实值。"""
    px = max(6, int(round(size_pt * 96.0 / 72.0)))
    return _font(size_pt, font_name), (size_pt * 96.0 / 72.0) / float(px)


def text_extent_mm(text, size_pt, width_mm, font_name="微软雅黑"):
    """按字体真实度量估算折行后的文字高度（mm）。

    返回 (需要的总高 mm, 折行行数)。

    为什么需要它：Visio 的 Shape 对象 **没有 TextBounds 属性/方法**
    （早绑定、晚绑定都取不到），无法直接问 Visio「这段字占多高」。
    因此改用 PIL 读字体文件，按框宽逐字试排来估算。
    """
    f, sc = _font_scale(size_pt, font_name)
    wpx = max(4.0, width_mm / 25.4 * 96.0)
    lines = 0
    for para in text.split("\n"):
        if not para:
            lines += 1
            continue
        cur, n = "", 1
        for ch in para:
            if cur and f.getlength(cur + ch) * sc > wpx:
                n += 1
                cur = ch
            else:
                cur += ch
        lines += n
    asc, desc = f.getmetrics()
    return lines * (asc + desc) * sc / 96.0 * 25.4, lines


def _poly_dist(pts, x, y):
    """点到折线的最短距离（页面坐标）。"""
    best = 1e18
    for i in range(len(pts) - 1):
        x1, y1 = pts[i]
        x2, y2 = pts[i + 1]
        dx, dy = x2 - x1, y2 - y1
        L2 = dx * dx + dy * dy
        t = 0.0 if L2 <= 1e-9 else max(0.0, min(1.0, ((x - x1) * dx + (y - y1) * dy) / L2))
        px, py = x1 + t * dx, y1 + t * dy
        d = ((x - px) ** 2 + (y - py) ** 2) ** 0.5
        if d < best:
            best = d
    return best


def _poly_near(pts, x, y):
    """点到折线的**最近点**信息：返回 `(距离, 段号, 最近点, 段单位方向)`。

    与 `_poly_dist` 同一套投影算法，只把中间量一起返回（页面坐标，Y 向下）。
    段方向用于「标签贴线归一」：文字要落在最近点的法向一侧。
    """
    best = None
    for i in range(len(pts) - 1):
        x1, y1 = pts[i]
        x2, y2 = pts[i + 1]
        dx, dy = x2 - x1, y2 - y1
        L2 = dx * dx + dy * dy
        if L2 <= 1e-9:
            continue
        t = max(0.0, min(1.0, ((x - x1) * dx + (y - y1) * dy) / L2))
        px, py = x1 + t * dx, y1 + t * dy
        d = ((x - px) ** 2 + (y - py) ** 2) ** 0.5
        if best is None or d < best[0]:
            best = (d, i, (px, py), (dx / L2 ** 0.5, dy / L2 ** 0.5))
    return best


def _fix_emf_size(path, W, H):
    """按 **rclBounds** 重写 EMF 的 mm 映射（★ RDP 会话必需；rclBounds 保持原样）。

    RDP 会话下 Visio 取不到屏幕物理尺寸，EMF 头里 szlMillimeters 退化为 1000×1000，
    于是 rclFrame（0.01 mm）按**非等比**换算写出 —— 宽高比失真（实测 图 N5 由 0.962 变
    0.524、图 9-1 由 1.41 变 0.74）。而 **rclBounds（设备像素 = 内容边界）本身始终正确**。
    排版侧只用 **rclFrame 的宽高比**定图幅，故这里只按 rclBounds 重写
    rclFrame / szlDevice / szlMillimeters（统一 K px/mm），**绝不改 rclBounds** ——
    保证图幅宽高比 = 内容宽高比，图不被拉伸。
    判据：szlDevice 与 szlMillimeters 的宽高比失配（正常会话二者一致）；未命中则原样返回。
    """
    import struct
    K = 4.62                    # Visio EMF 导出 px/mm（取自正常会话 szlDevice/szlMillimeters≈4.6295）
    with open(path, "rb") as f:
        buf = bytearray(f.read())
    if len(buf) < 88:
        return
    dev = struct.unpack_from("<ii", buf, 72)
    mil = struct.unpack_from("<ii", buf, 80)
    if dev[0] and dev[1] and mil[0] and mil[1]:
        dr, mr = dev[0] / float(dev[1]), mil[0] / float(mil[1])
        if abs(dr - mr) / dr < 0.05:
            return                      # 正常会话：rclFrame 已正确，不动
    rb = struct.unpack_from("<iiii", buf, 8)          # rclBounds（内容边界，像素）
    bw, bh = rb[2] - rb[0], rb[3] - rb[1]
    if bw <= 0 or bh <= 0:
        return
    mmw, mmh = bw / float(K), bh / float(K)
    struct.pack_into("<iiii", buf, 24, 0, 0, int(round(mmw * 100)), int(round(mmh * 100)))  # rclFrame
    struct.pack_into("<ii", buf, 72, bw, bh)                                              # szlDevice
    struct.pack_into("<ii", buf, 80, int(round(mmw)), int(round(mmh)))                     # szlMillimeters
    with open(path, "wb") as f:
        f.write(buf)


# 形状段索引（实测，Visio 2021）：Geometry=10 / ConnectionPts=7
GEOM_SEC = 10


class VDraw(object):
    """一个 Visio 绘图会话 = 一个页面。"""
    default_auto_merge = True                   # 子类/脚本可改；见 __init__
    def __init__(self, w=300.0, h=170.0, visible=False, theme="tech-cn"):
        self.W, self.H = float(w), float(h)
        self.t = _theme.get(theme)
        self.font = self.t.font
        self.app = win32.DispatchEx("Visio.Application")  # DispatchEx：独立实例，避免附加到既有实例时 COM 失效
        try:
            self.app.Visible = 1 if visible else 0
        except Exception:
            pass
        try:
            self.app.AlertResponse = 7          # 抑制「是否保存」等弹窗
        except Exception:
            pass
        self.doc = self.app.Documents.Add("")
        self.pg = self.doc.Pages.Item(1)
        self._texts = []
        self._n_dyn = 0                         # 动态路由连接线计数
        self._cps = {}                          # (形状ID, 行号) -> (fx, fy)，供出线方向自检
        self._link_warns = []                   # 连线方向/穿越自检发现的问题
        self._merged = False                    # merge_texts() 幂等标记
        self._merged_stat = None
        # save()/export() 前自动并入文字（可用环境变量 VISIO_MERGE=0 关闭，
        # 或用类属性 VDraw.default_auto_merge=False 关闭 → 便于出「未合并」对照图）
        env = os.environ.get("VISIO_MERGE", "")
        self.auto_merge = (env != "0") if env else VDraw.default_auto_merge
        self._page(w, h)

    # ---------------------------------------------------------- 基础
    def _page(self, w, h):
        ps = self.pg.PageSheet
        self._f(ps, "PageWidth", "%g mm" % w)
        self._f(ps, "PageHeight", "%g mm" % h)
        for c in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
            self._f(ps, c, "0 mm")

    @staticmethod
    def _f(obj, cell, formula):
        """安全写 CellsU：个别属性在某些 Visio 版本上不存在，忽略即可。"""
        try:
            obj.CellsU(cell).FormulaU = formula
        except Exception:
            pass

    def _blank(self):
        return self.pg.DrawRectangle(0, 0, 1, 1)

    def _place(self, s, x, ytop, w, h):
        self._f(s, "PinX", "%g mm" % (x + w / 2.0))
        self._f(s, "PinY", "%g mm" % (self.H - ytop - h / 2.0))
        self._f(s, "Width", "%g mm" % w)
        self._f(s, "Height", "%g mm" % h)

    def _skin(self, s, fill="white", line="black", lw="0.75 pt", lpat=1,
              rounding="0 mm", color="black", font=None, size=8.5,
              align=1, valign=1):
        t = self.t
        fill = None if fill is None else t.c(fill, fill)
        line = t.c(line, line)
        color = t.c(color, color)
        self._f(s, "FillPattern", "0" if fill is None else "1")
        if fill is not None:
            self._f(s, "FillForegnd", fill)
        self._f(s, "LinePattern", str(lpat))
        if lpat:
            self._f(s, "LineColor", line)
            self._f(s, "LineWeight", lw)
        self._f(s, "Rounding", rounding)
        self._f(s, "Shadows", "0")
        self._f(s, "Char.Font", '"%s"' % (font or self.font))
        self._f(s, "Char.Color", color)
        self._f(s, "Char.Size", "%g pt" % size)
        self._f(s, "Para.HorzAlign", str(align))     # 0左 1中 2右
        self._f(s, "VerticalAlign", str(valign))     # 0上 1中 2下
        self._f(s, "LeftMargin", "%g mm" % MARGIN_MM)
        self._f(s, "RightMargin", "%g mm" % MARGIN_MM)
        self._f(s, "TopMargin", "0.6 mm")
        self._f(s, "BottomMargin", "0.6 mm")

    # ---------------------------------------------------------- 图元
    def rect(self, x, ytop, w, h, **kw):
        """矩形。kw 见 _skin()，颜色可传主题字段名（如 'accent_fill'）。"""
        s = self._blank()
        self._place(s, x, ytop, w, h)
        self._skin(s, **kw)
        return s

    def text(self, x, ytop, w, h, content, size=8.5, color="black", align=1,
             valign=1, font=None):
        """无边框文本框。会登记到溢出自检清单。"""
        s = self._blank()
        self._place(s, x, ytop, w, h)
        self._skin(s, fill=None, lpat=0, color=color, font=font,
                   size=size, align=align, valign=valign)
        s.Text = content
        self._texts.append((content, size, w, h, font or self.font))
        return s

    def line(self, x1, y1, x2, y2, arrow=4, lpat=1, color="black",
             lw="0.75 pt", begin_arrow=0):
        """直线。arrow=4 为实心箭头，0 为无箭头；lpat: 1实线 2虚线 4点划线。"""
        s = self.pg.DrawLine(0, 0, 1, 1)
        self._f(s, "BeginX", "%g mm" % x1)
        self._f(s, "BeginY", "%g mm" % (self.H - y1))
        self._f(s, "EndX", "%g mm" % x2)
        self._f(s, "EndY", "%g mm" % (self.H - y2))
        self._f(s, "EndArrow", str(arrow))
        self._f(s, "BeginArrow", str(begin_arrow))
        self._f(s, "LinePattern", str(lpat))
        self._f(s, "LineColor", self.t.c(color, color))
        self._f(s, "LineWeight", lw)
        return s

    def poly(self, pts, arrow=4, lpat=1, color="black", lw="0.75 pt"):
        """折线：pts=[(x,y),...]，**只有最后一段带箭头**（避免每个拐点都是箭头）。

        ★ 这是「多条互不相连的线段」画法：不随方块移动、无法在 Visio 里整体编辑。
        需要与被连接方块联动的场合请改用 connect()（单条动态连接线 + 胶水）。
        """
        n = len(pts) - 1
        for i in range(n):
            self.line(pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1],
                      arrow=(arrow if i == n - 1 else 0),
                      lpat=lpat, color=color, lw=lw)

    # ---------------------------------------------------------- 连接线
    # ★ fy 自上边起算（本库统一约定：左上原点、Y 向下）：fy=0 上边 / fy=1 下边。
    #   Visio 连接点细胞是**局部坐标（Y 向上、原点在形状左下角）**，故内部写成
    #   =Height*(1-fy) —— 2026-10-12 实测踩坑：直接写 =Height*fy 会让 top/bottom 整体颠倒
    #   （竖线从「上框顶边」贯穿到「下框底边」）。DirX/DirY 同为局部坐标：+Y 向上。
    _SIDE_CP = {"top": (0.5, 0.0, 0, 1), "bottom": (0.5, 1.0, 0, -1),
                "left": (0.0, 0.5, -1, 0), "right": (1.0, 0.5, 1, 0)}
    # 四向中点的朝外法线（本库页坐标：Y 向下）——出线方向自检用
    _SIDE_N = {"left": (-1, 0), "right": (1, 0), "top": (0, -1), "bottom": (0, 1)}

    @classmethod
    def _cp_side(cls, fx, fy):
        """连接点比例 (fx,fy) → 'left'/'right'/'top'/'bottom'；非四向中点返回 None。"""
        if abs(fx) < 1e-6:
            return "left"
        if abs(fx - 1.0) < 1e-6:
            return "right"
        if abs(fy) < 1e-6:
            return "top"
        if abs(fy - 1.0) < 1e-6:
            return "bottom"
        return None

    def add_cp(self, s, fx=0.5, fy=0.5, dirx=0, diry=0):
        """给形状加一个连接点，位置 = 形状外框比例 (fx·Width, fy·Height)。

        fx 自左边缘起算（0 左 / 1 右）；**fy 自上边缘起算**（0 上 / 1 下 / 0.5 中），
        与本库公开坐标约定一致；写成公式 → 方块缩放时连接点自动跟随。
        dirx/diry：出线方向（Visio 局部坐标，+Y 向上）。
        返回 Connections 段行号（0 起），供 connect() 的 ia/ib 引用。
        """
        row = s.AddRow(7, -2, 0)          # visSectionConnectionPts=7 / visRowLast=-2 / visTagDefault=0
        n = row + 1
        self._f(s, "Connections.X%d" % n, "=Width*%.10g" % float(fx))
        self._f(s, "Connections.Y%d" % n, "=Height*(1-%.10g)" % float(fy))
        if dirx or diry:
            # DirX/DirY 必须走 CellsSRC：Connections 行里 CellsU("Connections.DirX1")
            # 在部分 Visio 版本上直接报错（Unexpected end of file），写不进去。
            # ConnectionPts 行列号：0=X 1=Y 2=DirX 3=DirY。
            try:
                s.CellsSRC(7, row, 2).FormulaU = "%g" % float(dirx)
                s.CellsSRC(7, row, 3).FormulaU = "%g" % float(diry)
            except Exception:
                pass
        # 记下 (fx, fy)，供 connect() 的「出线方向自检」判断这一端朝哪侧
        try:
            self._cps[(s.ID, row)] = (float(fx), float(fy))
        except Exception:
            pass
        return row

    def side_cps(self, s, sides="top bottom left right"):
        """批量加四向（或指定若干向）连接点，返回 {方向: 行号}。

        出线方向按朝向写入 DirX/DirY，动态连接线的直角路由会更干净
        （否则 Visio 自行判定，容易绕远）。例：
            cp = d.side_cps(s, "top bottom")     # 只要上/下
            d.connect(a, cp["right"], b, cp["left"], kind="elbow")
        """
        out = {}
        for k in sides.split():
            fx, fy, dx, dy = self._SIDE_CP[k]
            out[k] = self.add_cp(s, fx, fy, dirx=dx, diry=dy)
        return out

    def connect(self, sa, ia, sb, ib, arrow=4, lpat=1, color="black",
                lw="0.75 pt", via=None, label=None, label_size=7,
                label_color=None, label_align=1, label_pos=None,
                label_at=None, label_gap=2.2, label_side=None):
        """单条动态连接线：两端胶到 sa[ia] / sb[ib] 指定的连接点。

        ia/ib 为 add_cp() 返回的行号；传 None 表示胶到形状中心（PinX/PinY）；
        传 `(x, y)` 页坐标元组表示该端**不胶水**、固定在给定点（用于指向图外 /
        汇入总线的箭头）。首/尾点可省略，自动按端点补齐。
        via:  **路径点**（页坐标 mm，左上原点、Y 向下），如
                [(x1,y1),(k1x,k1y),...,(x2,y2)]
              首/尾点可省略，自动按胶水端点补齐。
              · via=None（默认）→ 直线连接（Begin 到 End）；
              · via 给 3 个点以上 → 直角拆线 / 折线，按给定折点走。
        路径写进几何行（单条形状），端点胶在连接点上：拖动方块时端点跟随；
        在 Visio 中整体选中、改线型、改端点都只针对这一个形状。
        ★ 不用 Page.Layout() 重排：它会把已手工排版的方块重新分布（实测把
        管线框摊到页外）。几何段索引 = 10，行命名 Geometry1.X{n}/Y{n}（IX=n）。

        label / label_size / label_color / label_align / label_pos：
          给连接线加**自带文本**（写在 1-D 形状的 Text 上，**不是独立文本框**）。
          ★ 推荐做法：文本随连接线移动，**选中时只显示端点手柄、不出现矩形框**
            （独立文本框选中时会画出自身外框，观感差、也易被误当成图形元素）。
          定位（★ 一律推荐用「沿线相对定位」，**不要**手写绝对坐标）：
            label_at  = None|float|(段号,比例)   沿线定位：
                        None → 自动取**最长一段的中点**（最稳，默认）
                        float → 按整条折线**弧长比例**取点（0=起点，1=终点）
                        (i, f) → 第 i 段（0 起）内按比例 f 取点
            label_gap = 文字中心到线段的垂距 mm（默认 2.2）
            label_side= None|'auto'|'up'|'down'|'left'|'right'|'on'
                        线段朝哪侧偏：None/'auto' 按线段走向自选
                        （横线段→上方，竖线段→左侧）；'on' = 压在线上
            label_pos = (x,y) 绝对页坐标（**不随线移动**，改布局后会离老远，
                        仅确需精确定位时用；用后务必肉眼看一遍成品）。
        """
        def build():
            c = self.pg.DrawLine(0, 0, 1, 1)      # 1-D：自带 Begin/End 细胞可胶
            for cell, shp, idx in (("Begin", sa, ia), ("End", sb, ib)):
                if isinstance(idx, tuple):        # (x, y)：自由端，固定在该点
                    self._f(c, cell + "X", "%g mm" % float(idx[0]))
                    self._f(c, cell + "Y", "%g mm" % (self.H - float(idx[1])))
                elif idx is None:
                    self._glue(c, cell + "X", shp, "PinX")
                    self._glue(c, cell + "Y", shp, "PinY")
                else:
                    self._glue(c, cell + "X", shp, "Connections.X%d" % (idx + 1))
                    self._glue(c, cell + "Y", shp, "Connections.Y%d" % (idx + 1))
            self._f(c, "EndArrow", str(arrow))
            self._f(c, "BeginArrow", "0")
            self._f(c, "LinePattern", str(lpat))
            self._f(c, "LineColor", self.t.c(color, color))
            self._f(c, "LineWeight", lw)
            return c

        c = build()
        free_end = isinstance(ia, tuple) or isinstance(ib, tuple)
        if free_end:
            # ★ 有一端是**自由点**（不胶水）的连接线一律不加 ObjType：
            #   Visio 对自由端的 1-D 形状会施加 0.25 in 最小尺寸夹取并重排几何
            #   （图 N10 时间刻度线实测被缩到 0.475 mm、图 N8 引出线被加出斜段）。
            self._author_path(c, via, objtype=False)
        elif not self._author_path(c, via, objtype=True):
            # ★ 自愈回退：个别**很短**的连接线（长度 < 0.25 in）或胶到**另一条
            #   连接线**上的线，仍可能被 Visio 的 1-D 最小尺寸夹取改坏 →
            #   删除重建，退化为「不加 ObjType」的老做法。
            #   代价：该条线选中时仍显示包围盒；但**图形本身正确**优先。
            try:
                c.Delete()
            except Exception:
                pass
            c = build()
            self._author_path(c, via, objtype=False)
        if label:
            self._set_conn_label(c, label, size=label_size,
                                 color=(label_color or color),
                                 align=label_align, pos=label_pos,
                                 at=label_at, gap=label_gap,
                                 side=label_side)
        self._check_end_dirs(c, sa, ia, sb, ib, via)
        return c

    def _check_end_dirs(self, c, sa, ia, sb, ib, via):
        """端点出线方向自检：**不得从所胶连接点的反方向出发**。

        ★ 为什么必须查（2026-10-17 项目实测）：`connect(a, cp["left"], b, cp["left"])`
          这种写法在 b 位于 a 右侧时，连线会从 a 的**左边**出发横穿整个 a。
          a 有底色 → 线被压在框下面、只在缺口露一截，成品看着像「从框左边划过来」，
          肉眼很难发现，源脚本也读不出错。
        判据：该端连接点的朝外法线 · 首段方向 < 0 → 首段指向框内，必穿框。
        只查**四向中点**（自定义连接点位置无「朝向」语义，跳过）；自由点/中心胶水跳过。
        """
        try:
            bx = c.CellsU("BeginX").Result("mm")
            by = c.CellsU("BeginY").Result("mm")
            ex = c.CellsU("EndX").Result("mm")
            ey = c.CellsU("EndY").Result("mm")
        except Exception:
            return
        # ★ 一律换成本库页坐标（左上原点、Y 向下）再比对：_SIDE_N 用的就是这套
        #   （BeginY/EndY 是 Visio 的 Y 向上坐标，直接用会让上下两侧的正负号反掉）
        pts = [(bx, self.H - by)]
        if via:
            pts += [(p[0], p[1]) for p in via[1:-1]]
        pts.append((ex, self.H - ey))
        if len(pts) < 2:
            return
        for tag, shp, idx, p0, p1 in (("Begin", sa, ia, pts[0], pts[1]),
                                      ("End", sb, ib, pts[-1], pts[-2])):
            if not isinstance(idx, int):
                continue
            info = self._cps.get((shp.ID, idx))
            if not info:
                continue
            side = self._cp_side(*info)
            if not side:
                continue
            dx, dy = p1[0] - p0[0], p1[1] - p0[1]
            L = (dx * dx + dy * dy) ** 0.5
            if L < 1e-6:
                continue
            nrm = self._SIDE_N[side]
            if (dx * nrm[0] + dy * nrm[1]) / L >= -0.5:
                continue
            msg = ("[连线方向] %s 端胶在形状#%s 的 %s 侧，首段却朝反方向 → 连线会横穿该形状。"
                   "改用朝向目标的那一侧（例：connect(a, cp[\"right\"], b, cp[\"left\"])）。"
                   % (tag, shp.ID, side))
            print("  " + msg)
            self._link_warns.append(msg)

    @staticmethod
    def _glue(conn, cell, shp, target):
        """把连接线的 Begin/End 细胞胶到目标形状的某个细胞上。"""
        conn.CellsU(cell).GlueTo(shp.CellsU(target))

    def _author_path(self, conn, via, objtype=True):
        """把路径写进连接线的几何行（via=None 时为 Begin→End 直线）。

        返回 True = 已按「可布线 1-D 形状」写好并校验通过；
        返回 False = Visio 的 1-D 最小尺寸夹取改坏了几何，请上层**删除重建**
        （objtype=False 再调一次即退化为老做法）。

        局部坐标 = 相对 Begin 点的偏移。★ 同时还必须满足 Visio 规范：
        **几何行必须落在局部坐标 [0,Width]×[0,Height] 内**（原点在左下、Y 向上）。
        折线向下/向左走时偏移为负，若不平移，Width/Height 与实际几何范围分离，
        形状外接框（= Visio 里的选中框）就落到别处 —— 用户选中折线时会在
        非折线区域看到一块透明空白矩形（2026-10-12 实测报告）。
        故这里把几何整体平移到左下角（dx0/dy0）。

        ★★ 「选中时沿包围盒出现一圈矩形外框」的**真正根因 = `ObjType`**
          （2026-10-13 单变量 A/B 逐项二分定位，9 组对照）：
          `Misc.ObjType` 默认 0（由 Visio 自行判断）→ 形状被按普通 2-D 形状渲染，
          选中即画出包围盒；原生 Dynamic connector 是 `ObjType = 2`
          （`visLOFlagsRoutable`，「可布线，必须是 1-D 形状」，见 MS 文档
          ObjType Cell, Miscellaneous Section）→ 选中只显示端点圆点 + 折点小方块。
          ⚠ 与 ShapeTransform 无关：`LocPin=包围盒中心` 与 `LocPin=包围盒角`
          两种写法**都出矩形**（早期判读有误，已由放大复核推翻）。

        ★★ 对策与副作用（**顺序敏感，务必照抄**）：
          先按包围盒中心把变换/几何写对 → `ConFixedCode=3` → 才设 `ObjType=2`
          → 最后按**回读的真实局部原点**重写几何。
          · 设 `ObjType=2` 会触发 Visio 布线引擎接管：把 ShapeTransform 重写成
            原生连接线约定 `PinX=GUARD((BeginX+EndX)/2)`、`LocPinX=GUARD(Width*0.5)`、
            `Width=GUARD(EndX-BeginX)`、`Height=GUARD(EndY-BeginY)`，并**丢弃自写
            几何**，故必须在**其后**重写几何。
          · `ConFixedCode` 必须**先**设：否则 Visio 在随后的重算里再次重布
            （实测同一横线得 6 点自动路由；先锁布线则得精确 4 点）。
          · Visio 把不足 0.25 in 的边**夹取到 ±0.25 in**（横线 Height、竖线 Width、
            零长线两向）→ 局部原点 (PinX−LocPinX, PinY−LocPinY) **≠ Begin 点**
            （实测横线偏 3.175 mm）→ 必须**回读**原点：`X = 页x − ox`、`Y = 页y − oy`。
          几何在重算 / 存盘 / 重开后均保持。
          · ★ 不足 0.25 in 的**短形状**会被夹取改坏（图 N10 时间刻度线实测被
            缩到 0.475 mm）→ 本方法**逐点校验**绝对坐标，不合格即返回 False，
            由 `connect()` 删除重建并退化为「不加 ObjType」（图形正确优先）。

        ★ 下游 `_line_page_pts()` 与标签写入都取 `PinX − LocPinX` 作局部原点，
          与上面回读的原点一致，故无需改动。

        行结构：RowCount(10) 含 component 行；路径行 IX=1..n，
        第 1 行 MoveTo，其余 LineTo（AddRow 的 tag=139 = visTagLineTo）。
        """
        bx = conn.CellsU("BeginX").Result("mm")
        by = conn.CellsU("BeginY").Result("mm")
        ex = conn.CellsU("EndX").Result("mm")
        ey = conn.CellsU("EndY").Result("mm")
        if via is None:
            loc = [(0.0, 0.0), (ex - bx, ey - by)]
        else:
            loc = [(p[0] - bx, (self.H - p[1]) - by) for p in via]
            loc[0] = (0.0, 0.0)
            loc[-1] = (ex - bx, ey - by)
        xs = [p[0] for p in loc]
        ys = [p[1] for p in loc]
        dx0, dy0 = min(xs), min(ys)          # 几何左下角（dx0 常为 0、dy0 常 ≤0）
        w = max(xs) - dx0
        h = max(ys) - dy0
        # ① 先按「包围盒中心」变换 + 平移几何 —— 保证设 ObjType 之前形状已正确成形
        self._f(conn, "LocPinX", "%.10g mm" % (w / 2.0))
        self._f(conn, "LocPinY", "%.10g mm" % (h / 2.0))
        self._f(conn, "PinX", "BeginX + %.10g mm" % (dx0 + w / 2.0))
        self._f(conn, "PinY", "BeginY + %.10g mm" % (dy0 + h / 2.0))
        self._f(conn, "Angle", "0")
        self._f(conn, "FlipX", "0")
        self._f(conn, "FlipY", "0")
        self._f(conn, "Width", "%g mm" % w)
        self._f(conn, "Height", "%g mm" % h)
        self._write_geom(conn, [(lx - dx0, ly - dy0) for lx, ly in loc])
        if not objtype:
            return True
        # ② 先**锁死布线**，再认作「可布线 1-D 形状」
        #    ★ 顺序关键：若先设 ObjType，Visio 会在随后的重算里**重布**并丢弃自写几何
        #      （实测同一横线：ObjType→写几何→ConFixedCode 得 6 点自动路由；
        #        ConFixedCode→ObjType→写几何 得精确 4 点）。
        self._f(conn, "ConFixedCode", "3")
        self._f(conn, "ObjType", "2")
        # ③ 按 Visio 接管后的**真实**局部原点重写几何
        #    ★ Visio 把 Width/Height 重写成 GUARD(EndX-BeginX)/GUARD(EndY-BeginY)，
        #      并把不足 0.25 in 的边**夹取到 ±0.25 in**（横线 Height、竖线 Width、
        #      零长线两向都夹）→ 局部原点 (PinX−LocPinX, PinY−LocPinY) **≠ Begin 点**
        #      （实测横线偏 3.175 mm）。故必须**回读**原点，不能按 Begin 推算。
        def origin():
            return (conn.CellsU("PinX").Result("mm") - conn.CellsU("LocPinX").Result("mm"),
                    conn.CellsU("PinY").Result("mm") - conn.CellsU("LocPinY").Result("mm"))

        ox, oy = origin()
        self._write_geom(conn, [(bx + lx - ox, by + ly - oy) for lx, ly in loc])
        # ④ 校验：Visio 的 1-D 最小尺寸夹取会把**短形状**的几何改坏（图 N10 刻度线
        #    实测被缩到 0.475 mm）→ 逐点比对绝对坐标，不合格则返回 False 交上层重建。
        try:
            ox2, oy2 = origin()
            n = conn.RowCount(GEOM_SEC) - 1
            got = [(ox2 + float(conn.CellsU("Geometry1.X%d" % (i + 1)).Result("mm")),
                    oy2 + float(conn.CellsU("Geometry1.Y%d" % (i + 1)).Result("mm")))
                   for i in range(n)]
        except Exception:
            return False
        if len(got) != len(loc):
            return False
        for (lx, ly), (gx, gy) in zip(loc, got):
            if abs((bx + lx) - gx) > 0.02 or abs((by + ly) - gy) > 0.02:
                return False
        return True

    def _write_geom(self, conn, pts):
        """把点列写入 Geometry1（IX=1..n，第 1 行 MoveTo、其余 LineTo）。"""
        n = len(pts)
        have = conn.RowCount(GEOM_SEC) - 1        # 去掉 component 行
        while have < n:
            conn.AddRow(GEOM_SEC, -2, 139)
            have += 1
        while have > n:
            conn.DeleteRow(GEOM_SEC, -2)
            have -= 1
        for i, (lx, ly) in enumerate(pts):
            r = i + 1
            conn.CellsU("Geometry1.X%d" % r).FormulaU = "%g mm" % lx
            conn.CellsU("Geometry1.Y%d" % r).FormulaU = "%g mm" % ly

    # ---------------------------------------------------------- 连接线自带文本
    @staticmethod
    def _label_anchor(pts, at=None):
        """在折线上取标签锚点：返回 **(锚点, 所在段方向)**（页坐标，Y 向下）。

        at=None        → **最长一段**的中点（最稳：线段最长、两侧最空）
        at=float       → 按整条折线**弧长比例**取点（0=起点，1=终点）
        at=(i, f)      → 第 i 段（0 起）内按比例 f 取点
        段方向用于算法向偏移；零长段跳过；段号越界回退到最长段中点。
        """
        if not pts:
            return (0.0, 0.0), (1.0, 0.0)
        if len(pts) < 2:
            return pts[0], (1.0, 0.0)
        segs = []
        for i in range(len(pts) - 1):
            a, b = pts[i], pts[i + 1]
            L = math.hypot(b[0] - a[0], b[1] - a[1])
            if L < 1e-6:
                continue
            segs.append((i, a, b, L))
        if not segs:
            return pts[0], (1.0, 0.0)
        if at is None:
            i, a, b, L = max(segs, key=lambda s: s[3])
            f = 0.5
        elif isinstance(at, tuple):
            i, f = int(at[0]), float(at[1])
            hit = [s for s in segs if s[0] == i]
            if hit:
                _, a, b, L = hit[0]
            else:
                _, a, b, L = max(segs, key=lambda s: s[3])
                f = 0.5
        else:
            total = sum(s[3] for s in segs)
            target = max(0.0, min(1.0, float(at))) * total
            acc, pick, f = 0.0, segs[-1], 1.0
            for s in segs:
                if acc + s[3] >= target - 1e-9:
                    pick, f = s, (target - acc) / s[3]
                    break
                acc += s[3]
            _, a, b, L = pick
        f = max(0.0, min(1.0, f))
        pt = (a[0] + (b[0] - a[0]) * f, a[1] + (b[1] - a[1]) * f)
        return pt, ((b[0] - a[0]) / L, (b[1] - a[1]) / L)

    @staticmethod
    def _perp_offset(pt, dv, gap, side):
        """把锚点沿线段的**法向**偏移 gap mm，得到文字块中心。

        side=None/'auto'：按线段走向自选 —— 横线段（|dx|≥|dy|）偏**上**、
        竖线段偏**左**（都是「读图时先看到」的一侧，且全图一致）。
        'up'/'down'/'left'/'right' 强制方向；'on' 或 gap=0 → 不偏移。
        """
        if not gap:
            return pt
        dx, dy = dv
        cands = ((-dy, dx), (dy, -dx))          # 线段两个法向
        if side in (None, "", "auto"):
            side = "up" if abs(dx) >= abs(dy) else "left"

        def _pick(axis, sign):
            for n in cands:
                if (n[0] if axis == "x" else n[1]) * sign > 1e-9:
                    return n
            return cands[0]

        n = {"up": _pick("y", -1), "down": _pick("y", 1),
             "left": _pick("x", -1), "right": _pick("x", 1),
             "on": (0.0, 0.0)}.get(side, cands[0])
        return (pt[0] + n[0] * gap, pt[1] + n[1] * gap)

    def _set_conn_label(self, c, text, size=7, color="gray", align=1,
                        pos=None, at=None, gap=2.2, side=None):
        """把标签写进 1-D 连接线**自带文本**（非独立文本框）。

        ★ 与 `_merge_line_labels` 同一套写入方式，但这里是**显式**调用 ——
          不受启发式的三道门槛（尺寸 / 距线距离 / 宿主分组）影响。
        定位 pos / at / gap / side 见 `connect()` 的 docstring：
          · pos=(x,y) 绝对坐标（**不随线移动**，改布局后会离老远）
          · 否则按 at 在折线上取锚点、再沿法向偏 gap mm
        ★ 四处必须显式设，否则文字位置/朝向会错：
          ① `TxtLocPinX/Y` = 半宽/半高 —— 1-D 形状默认 `TxtLocPinX = TxtWidth*0`
             （块左缘贴文本针）→ 文字右偏约 2mm；
          ② `TxtPinX/Y` 用**形状局部坐标**（原点 = PinX−LocPinX, PinY−LocPinY，Y 向上），
             不是 Begin 点（向左走的线上两者相差 LocPinX）；
          ③ `TxtAngle=0` —— 否则文字会随连接线角度一起旋转。

        ★★ 不要写 `Control.TextPosition` 段（2026-10-13 实测踩坑）：
          Visio 原生「动态连接线」确实带该命名行，但用 COM 按**形状局部坐标**写
          `AddNamedRow(9,'TextPosition',0)` + `CellsSRC(9,r,0/1)` 时，Visio 会把这两个
          值当**页面坐标**解释 → 文字跳到别处、并随连接线方向一起旋转（图 9-1 实例：
          「双 10GE / 组播 / 授时 / 授时信号」四标签全部错位且竖排）。
          只设 TxtPin/TxtLocPin/TxtAngle 即可正确水平落位，与 `_merge_line_labels` 一致。
        """
        c.Text = text
        self._f(c, "Char.Font", '"%s"' % self.font)
        self._f(c, "Char.Size", "%g pt" % size)
        self._f(c, "Char.Color", self.t.c(color, color))
        self._f(c, "Para.HorzAlign", str(int(align)))
        self._f(c, "VerticalAlign", "1")
        # ★★ 文本底块必须透明：主题/快速样式默认给 `TextBkgnd` 一个**白色**底块
        #   （V='#ffffff'），连线标签压在虚线/色块上时会**截断线条、显示成一个白框**
        #   —— 用户看到的就是这个「矩形框」（2026-10-13 实测：图 9-1「双 10GE」把
        #   竖虚线截成两段、「授时信号」把横虚线截断）。0 = 无底块。
        self._f(c, "TextBkgnd", "0")
        self._f(c, "TextBkgndTrans", "0")
        # 文本块尺寸按字体真实度量（PIL）；Visio 无 TextBounds 可问
        f, sc = _font_scale(size, self.font)
        tw = max(2.0, max(f.getlength(ln) for ln in text.split("\n"))
                 * sc / 96.0 * 25.4 + 4.5)
        nl = text.count("\n") + 1
        th = nl * self._lh(size) + 0.6
        # ★★ 连接线文本块**默认带左右页边距**（约 1.27mm/侧）→ 实际可用宽度比
        #   `TxtWidth` 小 ~2.5mm，两字标签会被挤成「一字一行」（2026-10-13 图 9-1
        #   「授时 / 组播」实测：TxtWidth 7.34mm 而 4.94mm 的文本仍换行）。
        #   对策：页边距清零 + 宽度留足余量。
        self._f(c, "LeftMargin", "0 mm")
        self._f(c, "RightMargin", "0 mm")
        self._f(c, "TxtWidth", "%g mm" % tw)
        self._f(c, "TxtHeight", "%g mm" % th)
        self._f(c, "TxtLocPinX", "%g mm" % (tw / 2.0))
        self._f(c, "TxtLocPinY", "%g mm" % (th / 2.0))
        if pos is not None:
            px, py = float(pos[0]), float(pos[1])
        else:
            # ★ 沿线相对定位（推荐）：锚点落在**线段上**，再朝线段某一侧偏 gap，
            #   于是「文字↔线」的距离恒定、且随连接线一起动。
            #   （旧兜底是取折线**包围盒中心** —— 肘形线的包围盒中心常常根本
            #    不在线上，文字会离老远；图 9-1 实测偏出 11~14mm。）
            pts, _ = self._line_page_pts(c)
            pt, dv = self._label_anchor(pts, at)
            px, py = self._perp_offset(pt, dv, gap, side)
        ox = c.CellsU("PinX").Result("mm") - c.CellsU("LocPinX").Result("mm")
        oy = c.CellsU("PinY").Result("mm") - c.CellsU("LocPinY").Result("mm")
        tx, ty = px - ox, (self.H - py) - oy
        self._f(c, "TxtPinX", "%g mm" % tx)
        self._f(c, "TxtPinY", "%g mm" % ty)
        self._f(c, "TxtAngle", "0")

    # ---------------------------------------------------------- 复合图元
    def box(self, x, ytop, w, h, title="", body="", tsize=8.5, bsize=7.0,
            fill="accent_fill", line="black", lw="0.8 pt", tcolor="accent",
            bcolor="dark", align=0, rounding="1.2 mm", pad=2.0):
        """标题 + 说明文字的方块；无 body 时标题垂直居中。
        返回底层矩形 shape，便于 add_cp()/connect() 挂连接线。"""
        s = self.rect(x, ytop, w, h, fill=fill, line=line, lw=lw, rounding=rounding)
        if body:
            self.text(x + pad, ytop + 1.3, w - 2 * pad, 6.2, title,
                      size=tsize, color=tcolor, align=align, valign=1)
            self.text(x + pad, ytop + 7.4, w - 2 * pad, h - 8.6, body,
                      size=bsize, color=bcolor, align=align, valign=0)
        else:
            self.text(x + pad, ytop, w - 2 * pad, h, title,
                      size=tsize, color=tcolor, align=align, valign=1)
        return s

    def step_box(self, x, ytop, w, h, title, body,
                 fill="accent_fill", line="black", lw="0.9 pt",
                 tcolor="accent", tsize=9.0, bsize=7.5):
        """流程图用的「标题 + 说明」框（标题居中，说明顶对齐）。
        返回底层矩形 shape，便于 add_cp()/connect() 挂连接线。"""
        s = self.rect(x, ytop, w, h, fill=fill, line=line, lw=lw,
                      rounding="1.5 mm")
        self.text(x + 1.5, ytop + 2.2, w - 3, 7.0, title,
                  size=tsize, color=tcolor, align=1, valign=1)
        self.text(x + 1.5, ytop + 10.5, w - 3, h - 13, body,
                  size=bsize, color="dark", align=1, valign=0)
        return s

    def lane(self, x, ytop, w, h, label="", lx=8.0, lw_=48.0,
             fill="lane", lfill="accent_fill", lcolor="accent", lsize=9.0):
        """泳道 / 分区：虚线外框 + 左侧标签块。"""
        self.rect(x, ytop, w, h, fill=fill, line="mgray", lw="0.75 pt", lpat=2)
        if label:
            self.rect(lx, ytop, lw_, h, fill=lfill, line=None, lpat=0)
            self.text(lx, ytop, lw_, h, label, size=lsize, color=lcolor,
                      align=1, valign=1)

    # ---------------------------------------------------------- 文字并入宿主形状
    # ★ 版式复原常数（2026-10-12 实测标定，见 pitfalls.md §18）
    #   行高 LH = 0.42911·字号 − 0.0514 mm（≈1.216×字号，残差 ≤0.1mm）
    #   首行墨迹 = 文字块顶 + 上边距 + 1.02 mm
    #   （★ 首版取 0.74 偏小：标定时 PNG 量化 1px=0.267mm 把均值拉低，
    #     合并回归实测均匀下移 1.0px → 校正为 1.02）
    #   ★★ 真块高 ≠ n×LH：Visio 做 va=1/2 居中时用的「文字块高」比行高累加多
    #     _BH_EXTRA（实测 8.5pt 与 9.0pt 均 ≈ +0.52mm；不补则整块下移 0.26mm）。
    #     va=0（顶对齐）不用块高 → 不受影响，故只补非 0 对齐。
    _LH_K, _LH_B, _INK = 0.42911, -0.0514, 1.02
    _BH_EXTRA = 0.52
    # 段间推进修正：Visio 实测「段→段首行墨迹」推进 A ≈ LH + 0.1~0.35（随字号/取整浮动）。
    # 用常数 _INK 反解 sp[] 时取一个折中值，实测可把多段正文的残差压到 ≤0.05mm。
    _ADV_BONUS = 0.15

    @classmethod
    def _lh(cls, size_pt):
        """Visio 单行行高（mm）。"""
        return cls._LH_K * float(size_pt) + cls._LH_B

    def merge_texts(self, tol=0.8, verbose=True):
        """把「落在可见矩形内」的纯文本框并入宿主矩形的**文字块**（多段富文本）。

        为什么值得做：合并前文字是独立形状，在 Visio 里拖动方块时**字会留在原地**；
        合并后字随框走，且形状数大幅下降（本项目 12 张图 584 → 310，−47%）。

        版式不变的做法：由原文本框的几何反解「合并后宿主的上边距 T」与「各段段前距」，
        使每段**首行墨迹**落回原来的位置：
            · 原框 valign=0 → 块顶 = 框顶 + 上边距；valign=1 → 块顶 = 框顶 + (框高 − 块高)/2
            · 块高 = 行数 × LH(字号) + _BH_EXTRA（★ 实测真块高比行高累加多 0.52mm）
            · T = 第 1 段墨迹 − 宿主框顶 − _INK
            · sp[i] = 墨迹[i] − 墨迹[i−1] − 行数[i−1] × (LH[i−1] + _ADV_BONUS)
        左右边距由原文本框的左右边界反推 → **折行结果与原来完全一致**。
        实测残差：标题/单行块 0.000mm，多段正文 ≤0.2mm（PNG 1px≈0.27mm，即 ≤1px）。

        跳过条件（保持原样，不强行合并）：宿主已有文字、同组文本框左右边界不一致、
        对齐方式不一致、字体不一致。

        返回 {'hosts': 合并的宿主数, 'texts': 并入的文本框数, 'skip': 跳过数, 'left': 未命中宿主数}
        """
        if getattr(self, "_merged", False):
            return self._merged_stat

        def _bbox(s):
            try:
                px = float(s.CellsU("PinX").Result("mm"))
                py = float(s.CellsU("PinY").Result("mm"))
                w = float(s.CellsU("Width").Result("mm"))
                h = float(s.CellsU("Height").Result("mm"))
            except Exception:
                return None
            # 返回页面坐标（左上原点、Y 向下）：(left, top, right, bottom)
            return (px - w / 2.0, self.H - (py + h / 2.0), px + w / 2.0, self.H - (py - h / 2.0))

        def _i(s, cell, dflt=0):
            try:
                return int(round(float(s.CellsU(cell).ResultIU)))
            except Exception:
                return dflt

        def _mm(s, cell, dflt=0.0):
            try:
                return float(s.CellsU(cell).Result("mm"))
            except Exception:
                return dflt

        def _pt(s, cell, dflt=8.5):
            try:
                return float(s.CellsU(cell).ResultIU) * 72.0
            except Exception:
                return dflt

        def _txt(s):
            try:
                return s.Text or ""
            except Exception:
                return ""

        # ---- 1. 分类：宿主（可见 2-D）／文本框（不可见但有字）／连接线（1-D，跳过）
        hosts, texts, nskip_shape = {}, {}, 0
        for i in range(1, self.pg.Shapes.Count + 1):
            s = self.pg.Shapes.Item(i)
            try:
                if s.OneD:
                    continue
                if s.Type == 2:                     # visTypeGroup：不递归，留给人工
                    nskip_shape += 1
                    continue
            except Exception:
                pass
            b = _bbox(s)
            if b is None:
                continue
            visible = _i(s, "LinePattern", 1) != 0 or _i(s, "FillPattern", 1) != 0
            t = _txt(s)
            if visible and not t.strip():
                hosts[s.ID] = (s, b)
            elif (not visible) and t.strip():
                texts[s.ID] = (s, b, t)

        # ---- 2. 每个文本框找「最小包含宿主」
        groups = {}
        left_over = 0
        left_items = []
        for tid, (s, b, t) in texts.items():
            cx, cy = (b[0] + b[2]) / 2.0, (b[1] + b[3]) / 2.0
            best, bestA = None, None
            for hid, (hs, hb) in hosts.items():
                if hb[0] - tol <= cx <= hb[2] + tol and hb[1] - tol <= cy <= hb[3] + tol:
                    a = (hb[2] - hb[0]) * (hb[3] - hb[1])
                    if bestA is None or a < bestA:
                        best, bestA = hid, a
            if best is None:
                left_over += 1
                left_items.append((s, b, t))
                if os.environ.get("VISIO_MERGE_DEBUG"):
                    print("  [DBG-无宿主] bbox=(%.2f,%.2f,%.2f,%.2f) va=%d size=%g  %s"
                          % (b[0], b[1], b[2], b[3], _i(s, "VerticalAlign", 1),
                             _pt(s, "Char.Size", 8.5), t.replace("\n", "|")[:30]))
                continue
            groups.setdefault(best, []).append((s, b, t))

        # ---- 3. 逐宿主合并
        n_host = n_text = n_skip = 0
        for hid, items in groups.items():
            host, hb = hosts[hid]
            if _txt(host).strip():
                n_skip += len(items)
                continue
            items.sort(key=lambda it: it[1][1])
            L, R = items[0][1][0], items[0][1][2]
            if any(abs(it[1][0] - L) > 0.4 or abs(it[1][2] - R) > 0.4 for it in items):
                n_skip += len(items)
                continue
            aligns = {_i(it[0], "Para.HorzAlign", 1) for it in items}
            fonts = {it[0].CellsU("Char.Font").FormulaU for it in items}
            if len(aligns) > 1 or len(fonts) > 1:
                n_skip += len(items)
                continue
            font = fonts.pop().strip('"')

            paras = []
            for s, b, t in items:
                size = _pt(s, "Char.Size", 8.5)
                lh = self._lh(size)
                lm, rm = _mm(s, "LeftMargin", MARGIN_MM), _mm(s, "RightMargin", MARGIN_MM)
                _, nline = text_extent_mm(t, size, (b[2] - b[0]) - lm - rm, font)
                bh = nline * lh + self._BH_EXTRA
                va = _i(s, "VerticalAlign", 1)
                tm = _mm(s, "TopMargin", 0.6)
                if va == 0:
                    btop = b[1] + tm
                elif va == 2:
                    btop = b[3] - tm - bh
                else:
                    btop = b[1] + (b[3] - b[1] - bh) / 2.0
                paras.append({"t": t, "size": size, "n": nline, "lh": lh,
                              "ink": btop + self._INK,
                              "color": s.CellsU("Char.Color").FormulaU,
                              "lm": lm, "rm": rm})

            T = max(0.0, paras[0]["ink"] - hb[1] - self._INK)
            sp = [0.0]
            for i in range(1, len(paras)):
                gap = (paras[i]["ink"] - paras[i - 1]["ink"]
                       - paras[i - 1]["n"] * (paras[i - 1]["lh"] + self._ADV_BONUS))
                sp.append(min(40.0, max(0.0, gap)))

            # ★ 逐段样式：**不能用 Characters.CharProps 设字号/颜色**
            #   类型库把 CharProps 的 newValue 声明为整数 → 8.5pt 被截成 8pt、
            #   颜色索引根本写不进去（实测 2026-10-12）。
            #   正解：先用**唯一哨兵字号** CharProps(7, 41+3i) 造出 run 边界，
            #   读回各行字号定位「段 → Character 段行号」，再用 CellsSRC 精修：
            #   Character 段（section 3）单元格 = 0 Font / 1 Color / 7 Size。
            host.Text = "\n".join(p["t"] for p in paras)
            self._f(host, "Char.Font", '"%s"' % font)
            self._f(host, "Char.Size", "%g pt" % paras[0]["size"])
            self._f(host, "Char.Color", paras[0]["color"])
            self._f(host, "Para.HorzAlign", str(aligns.pop()))
            self._f(host, "VerticalAlign", "0")
            self._f(host, "LeftMargin", "%g mm" % ((L - hb[0]) + paras[0]["lm"]))
            self._f(host, "RightMargin", "%g mm" % ((hb[2] - R) + paras[0]["rm"]))
            self._f(host, "TopMargin", "%g mm" % T)
            self._f(host, "BottomMargin", "0.6 mm")

            n_par = len(paras)
            sent = [41 + 3 * i for i in range(n_par)]
            c = host.Characters
            offs, off = [], 0
            for i, p in enumerate(paras):
                b0, e0 = off, off + len(p["t"])
                offs.append(b0)
                # 段尾换行也纳入本段范围（否则换行会单独成 run 并沿用错误字号，
                # 抬高本行行高 → 段间距被推大）
                c.Begin = b0
                c.End = e0 + (1 if i < n_par - 1 else 0)
                try:
                    c.CharProps(7, sent[i])
                except Exception:
                    pass
                off = e0 + 1

            rowmap = {}
            for r in range(0, n_par + 6):
                try:
                    fml = host.CellsSRC(3, r, 7).FormulaU
                except Exception:
                    break
                if not fml:
                    break
                mm_ = re.match(r"\s*([0-9.]+)\s*pt", fml)
                if mm_:
                    v = int(round(float(mm_.group(1))))
                    if v in sent:
                        rowmap[v] = r
            for i, p in enumerate(paras):
                r = rowmap.get(sent[i])
                if r is None:
                    continue
                for cell, val in ((0, '"%s"' % font), (1, p["color"]),
                                  (7, "%g pt" % p["size"])):
                    try:
                        host.CellsSRC(3, r, cell).FormulaU = val
                    except Exception:
                        pass
            # 附加 run（哨兵外的换行/空尾行）→ 沿用前一个已映射行的样式
            mapped = sorted(rowmap.values())
            for r in range(0, (max(mapped) + 3) if mapped else 0):
                if r in mapped:
                    continue
                prev = max([m for m in mapped if m < r], default=None)
                if prev is None:
                    continue
                for cell in (0, 1, 7):
                    try:
                        host.CellsSRC(3, r, cell).FormulaU = host.CellsSRC(3, prev, cell).FormulaU
                    except Exception:
                        pass

            # 段前距（★ 只取该段首字符：ParaProps 作用于范围内所有段落，
            #   取整段会把段前距也加到段内换行上，实测正文第 2 行被多推 3.9mm）
            for i in range(1, n_par):
                if sp[i] <= 0.05:
                    continue
                try:
                    c.Begin = offs[i]
                    c.End = offs[i] + 1
                    c.ParaProps(4, sp[i] * 72.0 / 25.4)   # mm → pt
                except Exception:
                    pass

            for s, b, t in items:
                try:
                    s.Delete()
                except Exception:
                    pass
            n_host += 1
            n_text += len(items)

        # ---- 4. 剩余文本框：落在某条 1-D 连接线上 → 并入连接线自带文本
        n_line = self._merge_line_labels(left_items) if left_items else 0

        self._merged = True
        self._merged_stat = {"hosts": n_host, "texts": n_text, "lines": n_line,
                             "skip": n_skip, "left": left_over - n_line,
                             "group": nskip_shape}
        if verbose:
            print("  [文字并入] 宿主 %d 个 / 并入文本框 %d 个 ｜ 连接线标签 %d 个"
                  " ｜ 跳过 %d（约束不一致） ｜ 无宿主 %d ｜ 组形状未处理 %d"
                  % (n_host, n_text, n_line, n_skip, left_over - n_line, nskip_shape))
        return self._merged_stat

    def _line_page_pts(self, s):
        """读一条 1-D 形状的几何折线，返回 **(页面坐标点列, 局部原点)**。

        ★★ 局部原点 ≠ Begin 点。`_author_path()` 会把几何整体平移到包围盒
        左下角（dx0/dy0），并把 Pin 定在包围盒中心、LocPin = 半宽/半高，于是
            局部原点 = (PinX − LocPinX, PinY − LocPinY) = 包围盒左下角
        （Visio 坐标，Y 向上）。老写法 `Begin + local` 只在 dx0=dy0=0
        （即向右/向下走的线）时恰好成立，向左/向上的线会整体错位
        —— 实测 图N9 的「上报」返程线（Begin=94、LocPinX=18）被判成
        距标签 9.86mm 而漏并，实为 5.00mm。

        ★ 行数用 `s.RowCount(10)` 界定：`s.RowCount("Geometry1")` 会抛
        `invalid literal for int()`；而对不存在的行
        `CellsU("Geometry1.X3").Result("mm")` **返回 0.0 而不报错**，
        故不能靠「读到 None 就停」截断（会读出一串 (0,0) 假点）。
        """
        def _num(cell):
            try:
                return float(s.CellsU(cell).Result("mm"))
            except Exception:
                return None

        pinx, piny = _num("PinX"), _num("PinY")
        lpx, lpy = _num("LocPinX"), _num("LocPinY")
        if None in (pinx, piny, lpx, lpy):
            return [], (0.0, 0.0)
        ox, oy = pinx - lpx, piny - lpy
        try:
            nrow = int(s.RowCount(GEOM_SEC))
        except Exception:
            nrow = 2
        pts = []
        for r in range(1, max(1, nrow)):        # 第 1 行是 component 行
            x, y = _num("Geometry1.X%d" % r), _num("Geometry1.Y%d" % r)
            if x is None or y is None:
                break
            pts.append((ox + x, self.H - (oy + y)))   # 页面坐标：左上原点、Y 向下
        return pts, (ox, oy)

    def _merge_line_labels(self, items, tol=6.0, max_w=44.0, max_h=12.0):
        """把「落在 1-D 连接线附近」的**小文本框**并入连接线**自带文本**。

        用 TxtPinX/TxtPinY 把文字钉回原来的中心（**形状局部坐标**，Y 向上），
        TxtAngle=0 保证文字始终水平；文本随连接线一起移动。
        只处理「标签尺寸」的框（宽 ≤ max_w、高 ≤ max_h）→ 不误吞标题/说明段。
        ★ 并入后做**贴线归一**：保持沿线位置与左右侧不变，只把「字形到线的
          距离」收敛到 LBL_GAP（0.8mm）—— 否则作者摆放的偏移会被原样保留，
          离线 3~5mm 而看着像独立说明（见 pitfalls §23）。
        """
        def _num(o, cell, dflt=None):
            try:
                return float(o.CellsU(cell).Result("mm"))
            except Exception:
                return dflt

        lines = []
        for i in range(1, self.pg.Shapes.Count + 1):
            s = self.pg.Shapes.Item(i)
            try:
                if not s.OneD:
                    continue
            except Exception:
                continue
            pts, org = self._line_page_pts(s)
            if len(pts) >= 2:
                lines.append((s, pts, org))
        if not lines:
            return 0

        n = 0
        for s, b, t in items:
            w, h = b[2] - b[0], b[3] - b[1]
            if w > max_w or h > max_h:      # 只并「标签尺寸」的小框
                continue
            cx, cy = (b[0] + b[2]) / 2.0, (b[1] + b[3]) / 2.0
            # ★ 2026-10-13 实测：把判据从「中心距线」放宽为「框内离线的最近点」
            #   会让**多个文本框并入同一条连接线**（`best.Text = t` 后者覆盖前者）
            #   → 图 8-1 的「输出文件复制回本机（回导）」整条丢失。故保持中心判据。
            best, borg, bd, bpts = None, None, None, None
            for ls, pts, org in lines:
                d = _poly_dist(pts, cx, cy)
                if bd is None or d < bd:
                    best, borg, bd, bpts = ls, org, d, pts
            if best is None or bd > tol:
                if os.environ.get("VISIO_MERGE_DEBUG"):
                    print("  [DBG-线标签未命中] %-16s c=(%.2f,%.2f) %.1fx%.1f 最近线距 %.2f (tol %.1f)"
                          % (t.replace("\n", "|")[:16], cx, cy, w, h,
                             -1 if bd is None else bd, tol))
                continue
            if os.environ.get("VISIO_MERGE_DEBUG"):
                print("  [DBG-线标签并入] %-16s c=(%.2f,%.2f) 距 %.2fmm"
                      % (t.replace("\n", "|")[:16], cx, cy, bd))
            # ★ 贴线归一（2026-10-13）：并进来的标签停在**作者摆放的位置**，
            #   实测普遍离线 3.5~5.5mm（fig_n8「生成/供」5.0/5.5、fig_n9 三条
            #   5.0、fig_n10 两条 3.5）→ 看着像一段独立说明、不像连接线标签。
            #   用户原话「连接线上的文本离连接线距离太远」。
            #   这里**保持「沿线位置」和「左右侧」不变**，只把**字形到线的
            #   距离**收敛到 LBL_GAP；原本就压在线上（距 ≈0）→ 按线段走向
            #   选默认侧（横→上、竖→左）。
            try:
                nd, si, npt, dv = _poly_near(bpts, cx, cy)
            except Exception:
                npt = None
            if npt is not None:
                try:
                    sz = float(s.CellsU("Char.Size").Result("pt"))
                except Exception:
                    sz = 7.0
                nl = t.count("\n") + 1
                gh = nl * self._lh(sz)                     # 字形总高
                try:
                    fnt, sc = _font_scale(sz, self.font)
                    gw = max(2.0, max(fnt.getlength(ln)
                                      for ln in t.split("\n")) * sc / 96.0 * 25.4)
                except Exception:
                    gw = max(2.0, w - 4.5)                 # 字形总宽（退化）
                if bd > 0.05:
                    ux, uy = (cx - npt[0]) / bd, (cy - npt[1]) / bd
                else:
                    p1 = self._perp_offset(npt, dv, 1.0, None)
                    ux, uy = p1[0] - npt[0], p1[1] - npt[1]
                # 轴对齐块沿 u 方向从中心到**近边**的距离（支撑函数）
                a, b = gw / 2.0, gh / 2.0
                den = math.sqrt((b * ux) ** 2 + (a * uy) ** 2)
                reach = (a * b / den) if den > 1e-9 else 0.0
                cx, cy = (npt[0] + ux * (LBL_GAP + reach),
                          npt[1] + uy * (LBL_GAP + reach))
                if os.environ.get("VISIO_MERGE_DEBUG"):
                    print("  [DBG-线标签贴线] %-16s 距 %.2f → %.2fmm（字形 %.1fx%.1f）"
                          % (t.replace("\n", "|")[:16], bd,
                             ((cx - npt[0]) ** 2 + (cy - npt[1]) ** 2) ** 0.5,
                             gw, gh))
            ox, oy = borg
            try:
                best.Text = t
            except Exception:
                continue
            self._f(best, "Char.Font", s.CellsU("Char.Font").FormulaU)
            self._f(best, "Char.Size", s.CellsU("Char.Size").FormulaU)
            self._f(best, "Char.Color", s.CellsU("Char.Color").FormulaU)
            try:
                self._f(best, "Para.HorzAlign", str(int(round(float(s.CellsU("Para.HorzAlign").ResultIU)))))
            except Exception:
                pass
            self._f(best, "VerticalAlign", "1")
            self._f(best, "TxtWidth", "%g mm" % w)
            self._f(best, "TxtHeight", "%g mm" % h)
            # ★ 1-D 形状的 TxtLocPinX 默认是 `TxtWidth*0`（块左缘贴文本针）→
            #   必须显式设成半宽/半高，块中心才落在 TxtPin 上（否则文字右偏 ~2mm）
            self._f(best, "TxtLocPinX", "%g mm" % (w / 2.0))
            self._f(best, "TxtLocPinY", "%g mm" % (h / 2.0))
            # ★ TxtPin 用**形状局部坐标**（原点 = PinX−LocPinX, Y 向上）→ 必须减
            #   局部原点，而不是 Begin 点（向左走的线上两者相差 LocPinX）。
            self._f(best, "TxtPinX", "%g mm" % (cx - ox))
            self._f(best, "TxtPinY", "%g mm" % ((self.H - cy) - oy))
            self._f(best, "TxtAngle", "0")
            try:
                s.Delete()
            except Exception:
                pass
            n += 1
        return n

    def h_arrow(self, x1, y, x2, **kw):
        """水平箭头（x1 -> x2）。"""
        self.line(x1, y, x2, y, **kw)

    def v_arrow(self, x, y1, y2, **kw):
        """垂直箭头（y1 -> y2）。"""
        self.line(x, y1, x, y2, **kw)

    def legend(self, x, y, w, items, lw_=26.0, gap=6.0, size=7.5,
               color="gray"):
        """图例行：items = [(说明文字, lpat 或 None), ...]。

        lpat=None 表示画实线样例（实心线），否则用给定线型。

        ★ 各条目在给定总宽 `w` 内**均匀分布**：整行自 `x` 起、到 `x+w` 止，
          每格宽 = (w − (n−1)·gap) / n，文字框宽 = 格宽 − 线样宽 − 1.5。
          （旧版**忽略 w**、文字框硬编码 46mm → 条目多时整行远远超出画布，
           导出按内容边界裁剪，右侧凭空多出一大片空白。见 pitfalls §22。）
          传 `w <= 0` 时退化为「线样 26 + 文字 46 + gap」的固定顺排。
        """
        n = len(items)
        if n == 0:
            return x
        if w and w > 0:
            cell = (float(w) - gap * (n - 1)) / n
        else:
            cell = lw_ + 1.5 + 46.0            # 旧固定顺排（整行宽 = n·cell + (n−1)·gap）
        tw = max(6.0, cell - lw_ - 1.5)
        cx = x
        for label, lpat in items:
            self.line(cx, y, cx + lw_, y, arrow=0,
                      lpat=(1 if lpat is None else lpat),
                      color=color, lw="0.9 pt")
            self.text(cx + lw_ + 1.5, y - 2.6, tw, 5.2, label,
                      size=size, color=color, align=0, valign=1)
            cx += cell + gap
        return cx - gap

    # ---------------------------------------------------------- 输出
    def reroute(self):
        """保留接口，**不再调用 Page.Layout()**。

        ★★ Page.Layout() 是「页面重排」：除了重算连接线路由，还会把已手工
        排版的方块重新分布（实测把管线框摊到页外）。本库所有走线一律由
        connect() 显式写几何行，不需要重排，故此处不再触发 Layout。
        """
        return

    def save(self, path):
        p = os.path.abspath(path)
        d = os.path.dirname(p)
        if d and not os.path.isdir(d):
            os.makedirs(d)
        if self.auto_merge:
            self.merge_texts()
        self.reroute()
        self.doc.SaveAs(p)
        return p

    def export(self, path):
        """按扩展名导出。Page.Export **按内容边界裁剪**（不含页边距），96 DPI。"""
        p = os.path.abspath(path)
        d = os.path.dirname(p)
        if d and not os.path.isdir(d):
            os.makedirs(d)
        if self.auto_merge:
            self.merge_texts()
        self.reroute()
        self.pg.Export(p)
        if p.lower().endswith(".emf"):
            _fix_emf_size(p, self.W, self.H)
        return p

    def export_all(self, base, strict=True):
        """一次导出三件套：base.vsdx / base.emf / base.png。返回**成功**的路径列表。

        ★ 旧版把导出异常吞掉只打印 → 调用方拿不到失败信号，会把「少导出一个」
          当成成功。现改为：三项**都会尝试**，但只要有失败，`strict=True`（默认）
          就抛 `RuntimeError` 列出失败项（`strict=False` 时仅打印、返回成功项）。
        """
        base = os.path.abspath(base)
        if base.lower().endswith((".vsdx", ".emf", ".png")):
            base = os.path.splitext(base)[0]
        out, fails = [], []
        for ext in (".vsdx", ".emf", ".png"):
            try:
                out.append(self.save(base + ext) if ext == ".vsdx"
                           else self.export(base + ext))
            except Exception as e:
                fails.append("%s%s : %s" % (base, ext, e))
                print("  [导出失败] %s%s : %s" % (base, ext, e))
        if fails and strict:
            raise RuntimeError("export_all 失败 %d/%d 项：\n  %s"
                               % (len(fails), 3, "\n  ".join(fails)))
        return out

    def close(self):
        for fn in (lambda: self.doc.Close(), lambda: self.app.Quit()):
            try:
                fn()
            except Exception:
                pass

    # ---------------------------------------------------------- 上下文管理
    def __enter__(self):
        """支持 `with VDraw(...) as d:` —— 退出时自动 close()，异常也不残留 VISIO.EXE。"""
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()
        return False                 # 不吞异常

    # ---------------------------------------------------------- 自检
    def report(self, tol=0.0):
        """返回可能溢出的文字清单 [(文字, 字号, 框宽, 框高, 行数, 需高)]。"""
        bad = []
        for txt, size, w, h, fname in self._texts:
            avail = w - 2 * MARGIN_MM
            need, lines = text_extent_mm(txt, size, avail, fname)
            if need > h + tol:
                bad.append((txt.replace("\n", " / ")[:36], size,
                            round(w, 1), round(h, 1), lines, round(need, 1)))
        return bad

    def check(self, verbose=True):
        """打印文字溢出自检 + 连线自检结果。返回溢出清单（空列表 = 通过）。"""
        bad = self.report()
        if verbose:
            if bad:
                print("  [文字溢出] %d 处" % len(bad))
                for t, sz, w, h, ln, need in bad:
                    print("    %-38s %gpt 框宽%.1f 框高%.1f → %d 行需 %.1fmm"
                          % (t, sz, w, h, ln, need))
            else:
                print("  [文字溢出] 无")
        cross = self.audit_links()
        if verbose:
            n1, n2 = len(self._link_warns), len(cross)
            if n1 or n2:
                print("  [连线自检] 出线方向 %d 处 ｜ 穿越实体框 %d 处（详见上方逐条提示）"
                      % (n1, n2))
            else:
                print("  [连线自检] 出线方向 / 穿越实体框 均无")
        return bad

    # ---------------------------------------------------------- 连线审计
    def _page_boxes(self):
        """页面上所有 2-D 实体框：[(ID, (x0,y0,x1,y1) 页坐标mm, FillPattern)]。"""
        out = []
        try:
            n = self.pg.Shapes.Count
        except Exception:
            return out
        for i in range(1, n + 1):
            try:
                s = self.pg.Shapes.Item(i)
                if s.OneD:
                    continue
                w = s.CellsU("Width").Result("mm")
                h = s.CellsU("Height").Result("mm")
                if w <= 1 or h <= 1:
                    continue
                x0 = s.CellsU("PinX").Result("mm") - s.CellsU("LocPinX").Result("mm")
                y1 = self.H - (s.CellsU("PinY").Result("mm") - s.CellsU("LocPinY").Result("mm"))
                # ★ 必须用 .Result(0)（visNone，取内部值）：传 "INT" 会报
                #   「Bad measurement unit.」→ 异常被吞 → 整个框被跳过、审计恒 0。
                fill = int(s.CellsU("FillPattern").Result(0))
                out.append((s.ID, (x0, y1 - h, x0 + w, y1), fill))
            except Exception:
                continue
        return out

    def _page_lines(self):
        """页面上所有 1-D 形状的折线点（页坐标 mm、Y 向下）。"""
        out = []
        try:
            n = self.pg.Shapes.Count
        except Exception:
            return out
        for i in range(1, n + 1):
            try:
                s = self.pg.Shapes.Item(i)
                if not s.OneD:
                    continue
                ox = s.CellsU("PinX").Result("mm") - s.CellsU("LocPinX").Result("mm")
                oy = s.CellsU("PinY").Result("mm") - s.CellsU("LocPinY").Result("mm")
                k = s.RowCount(GEOM_SEC) - 1
                pts = [(ox + s.CellsU("Geometry1.X%d" % (j + 1)).Result("mm"),
                        self.H - (oy + s.CellsU("Geometry1.Y%d" % (j + 1)).Result("mm")))
                       for j in range(k)]
                if len(pts) >= 2:
                    out.append((s.ID, pts))
            except Exception:
                continue
        return out

    @staticmethod
    def _seg_hits_rect(p, q, r, pad=0.25):
        """线段 pq 是否穿入矩形 r 的**内部**（Liang-Barsky，pad 收边 → 贴边不算）。"""
        x0, y0, x1, y1 = r[0] + pad, r[1] + pad, r[2] - pad, r[3] - pad
        if x1 <= x0 or y1 <= y0:
            return False
        dx, dy = q[0] - p[0], q[1] - p[1]
        t0, t1 = 0.0, 1.0
        for pp, qq in ((-dx, p[0] - x0), (dx, x1 - p[0]), (-dy, p[1] - y0), (dy, y1 - p[1])):
            if abs(pp) < 1e-12:
                if qq < 0:
                    return False
            else:
                t = qq / pp
                if pp < 0:
                    t0 = max(t0, t)
                else:
                    t1 = min(t1, t)
                if t0 > t1:
                    return False
        return t1 - t0 > 1e-9

    def audit_links(self, verbose=True):
        """整页连线审计：折线是否**穿入某个有底色实体框**内部。

        ★ 排除两类「穿过去也正常」的框：
          · 容器（泳道 / 系统边界框）——内部含 ≥2 个其它框的中心；
          · 无底色文本框（FillPattern=0）——通栏标题框，线从空白处经过不压字。
        返回问题清单；同时把「出线方向」的问题（connect() 时已记）一并打印。
        """
        boxes = self._page_boxes()
        page_area = max(1.0, self.W * self.H)
        kids = {}
        for bid, r, _f in boxes:
            # 子框只数「有底色」的：step_box 内部还有 1~2 个无底色文字框（保存前未 merge），
            # 若一并计数，每个 step_box 都会被误判成容器而漏查。
            kids[bid] = sum(1 for oid, o, g in boxes if oid != bid and g != 0 and
                            r[0] <= (o[0] + o[2]) / 2 <= r[2] and
                            r[1] <= (o[1] + o[3]) / 2 <= r[3])
        solid = []
        for bid, r, f in boxes:
            if f == 0:
                continue                           # 无底色文本框：线从空白处过不压字
            area = (r[2] - r[0]) * (r[3] - r[1])
            # 容器 = 大面积（≥18% 页面积，泳道/系统边界框）或内含 ≥2 个实体子框
            if area >= 0.18 * page_area or kids[bid] >= 2:
                continue
            solid.append((bid, r))
        bad = []
        for lid, pts in self._page_lines():
            for bid, r in solid:
                if all(r[0] <= p[0] <= r[2] and r[1] <= p[1] <= r[3] for p in pts):
                    continue                       # 容器：整条线都在里面
                if any(self._seg_hits_rect(p, q, r) for p, q in zip(pts[:-1], pts[1:])):
                    msg = "[连线穿越] 线#%s 穿入 形状#%s（%.0f×%.0f @%.1f,%.1f）内部" % (
                        lid, bid, r[2] - r[0], r[3] - r[1], r[0], r[1])
                    if verbose:
                        print("  " + msg)
                    bad.append(msg)
        return bad
