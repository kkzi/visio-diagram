# -*- coding: utf-8 -*-
"""
配色主题 —— 供 visio_lib 使用。

所有颜色均为 **Visio 公式字符串**（形如 "RGB(31,78,121)"），
直接写入 CellsU 的 FillForegnd / LineColor / Char.Color。
"""

import os

# 常见中文字体 -> 字体文件（用于文字溢出的字体度量自检）
FONT_FILES = {
    "微软雅黑": r"C:\Windows\Fonts\msyh.ttc",
    "Microsoft YaHei": r"C:\Windows\Fonts\msyh.ttc",
    "宋体": r"C:\Windows\Fonts\simsun.ttc",
    "SimSun": r"C:\Windows\Fonts\simsun.ttc",
    "黑体": r"C:\Windows\Fonts\simhei.ttf",
    "SimHei": r"C:\Windows\Fonts\simhei.ttf",
    "等线": r"C:\Windows\Fonts\Deng.ttf",
    "Arial": r"C:\Windows\Fonts\arial.ttf",
}


_FONT_WARNED = set()


def font_file(font_name, warn=True):
    """返回字体文件路径；找不到时回退到任意一个可用的系统字体。

    ★ 回退是**静默失真**：量字宽用的是回退字体，与 Visio 里实际渲染的字体
      不一致 → 折行行数 / 文字高度估算会偏，溢出可能漏报。
      故 `warn=True`（默认）时，对**每个缺失字体只警告一次**（走 stderr，
      不污染正常输出），提示溢出估算可能不准。
    """
    p = FONT_FILES.get(font_name)
    if p and os.path.isfile(p):
        return p
    for cand in FONT_FILES.values():
        if os.path.isfile(cand):
            if warn and font_name not in _FONT_WARNED:
                _FONT_WARNED.add(font_name)
                import sys as _sys
                print("  !! 字体 %r 无对应字体文件，已回退到 %s —— "
                      "文字溢出估算可能不准（如需精确请补 theme.FONT_FILES）"
                      % (font_name, os.path.basename(cand)), file=_sys.stderr)
            return cand
    return None


class Theme(object):
    """一套配色。字段用 self.__dict__ 展开，便于任意扩展。"""

    def __init__(self, name, font, **colors):
        self.name = name
        self.font = font
        self.__dict__.update(colors)

    def c(self, key, default=None):
        return getattr(self, key, default)

    def __repr__(self):
        return "<Theme %s font=%s>" % (self.name, self.font)


# ---------------------------------------------------------------- 内置主题

# 正式技术文档（白底 + 柔和色块 + 深灰边框，2026-10-10 按用户截图风格化）
TECH_CN = Theme(
    "tech-cn", "微软雅黑",
    black="RGB(62,63,63)",          # 图线主色（截图深灰 #3E3F3F）
    dark="RGB(38,38,38)",           # 正文说明文字
    gray="RGB(89,89,89)",           # 次要标注
    mgray="RGB(150,150,150)",       # 分隔框 / 虚线框
    lgray="RGB(242,242,242)",
    white="RGB(255,255,255)",       # 画布底色
    accent="RGB(56,71,103)",        # 标题 / 强调（深蓝灰）
    accent_fill="RGB(218,226,243)", # 主填充（截图 #DAE2F3 淡蓝紫）
    green_fill="RGB(226,239,217)",  # 次级填充（截图 #E2EFD9 淡绿）
    cream_fill="RGB(255,242,204)",  # 强调 / 外部块（截图 #FFF2CC 米黄）
    lane="RGB(242,246,252)",        # 泳道 / 分区底色（近白微蓝）
    muted_fill="RGB(226,239,217)",  # 外部对象 / 非本方模块 → 淡绿
)

# 纯黑白灰 —— 需要严格黑白印刷时用
MONO = Theme(
    "mono", "微软雅黑",
    black="RGB(0,0,0)",
    dark="RGB(51,51,51)",
    gray="RGB(102,102,102)",
    mgray="RGB(178,178,178)",
    lgray="RGB(238,238,238)",
    white="RGB(255,255,255)",
    accent="RGB(0,0,0)",
    accent_fill="RGB(221,221,221)",
    lane="RGB(250,250,250)",
    muted_fill="RGB(230,230,230)",
)

# 蓝色系 —— 汇报材料 / 幻灯片
BLUE = Theme(
    "blue", "微软雅黑",
    black="RGB(0,0,0)",
    dark="RGB(31,41,55)",
    gray="RGB(75,85,99)",
    mgray="RGB(156,163,175)",
    lgray="RGB(243,244,246)",
    white="RGB(255,255,255)",
    accent="RGB(21,94,169)",
    accent_fill="RGB(219,234,254)",
    lane="RGB(239,246,255)",
    muted_fill="RGB(229,231,235)",
)

THEMES = {t.name: t for t in (TECH_CN, MONO, BLUE)}


def get(name="tech-cn"):
    """按名字取主题；也可直接传入 Theme 实例。"""
    if isinstance(name, Theme):
        return name
    if name not in THEMES:
        raise KeyError("未知主题 %r，可选：%s" % (name, sorted(THEMES)))
    return THEMES[name]
