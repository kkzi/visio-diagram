# visio-diagram

用 Python + pywin32 驱动本机 Visio，程序化生成**原生可编辑**的矢量图，并嵌入 Word 文档。
面向中文工程技术文档配图（总体设计方案 / 技术方案 / 试验大纲 / 研究报告）。

```
产物三件套：*.vsdx（原生可编辑） · *.emf（矢量，供 Word/PPT） · *.png（预览）
```

## 能做什么

- **精确版式控制**：矩形 / 文本 / 泳道 / 图例 / 单条折线连接线（两端胶到连接点，挪框跟随）
- **文字溢出自检**：`TextBounds` 不存在 → 用 PIL 字体度量按框宽估算折行，出图前发现溢出
- **箭头对齐审计**：`audit_align.py` 查「箭头连到框但落点偏离边中点」的漂移；已定性为刻意
  的偏移可 `--add` 登记白名单（坐标 ±3mm 匹配、失效自动提示）
- **文字并入宿主**：保存前自动把「矩形内的文本框」并入矩形、把「贴线小标签」并入连接线，
  形状数省约一半、版式逐像素不变
- **连接线自带标签**：`connect(..., label=, label_at=, label_gap=, label_side=)` 沿线相对定位，
  随线移动
- **嵌入 Word 三条路线**：占位符替换（EMF 矢量）/ 手拼 OOXML / 原地换 OLE 对象
  （双击进 Visio 改图，含防图幅压扁、防页眉页脚丢失的实测配方）
- **文档自检**：`docx_audit.py`（静态 OOXML + Word 实测）、`selftest.py`（14 项回归断言）、
  `doc_consistency.py`（文档-代码一致性）

## 环境要求

- Windows + 桌面版 Visio（经 COM 驱动；PowerShell 无法实例化 COM，须走 Python）
- Python **3.9 ~ 3.13**（3.14+/alpha 预览版勿用：Pillow 的 `ImageFont` C 扩展会崩）
- `pip install pywin32 pillow`
- 中文字体（默认微软雅黑）

```bash
# 前置自检（Python 版本 / pywin32 / 字体度量一次查清）
python -c "import sys; sys.path.insert(0,'scripts'); import visio_lib; visio_lib.check_env()"
```

## 快速开始

```python
import sys
sys.path.insert(0, "scripts")
from visio_lib import VDraw

with VDraw(150, 108, theme="tech-cn") as d:      # 画布 mm，左上原点 / Y 向下
    d.lane(0, 0, 150, 74, "处理链路")
    d.step_box(6, 12, 42, 26, "① 环节一", "说明文字")
    d.h_arrow(48, 25, 60)
    d.check()                                     # 文字溢出自检
    d.export_all("out/图N1")                      # vsdx + emf + png
```

改完库先跑回归：

```bash
python scripts/selftest.py          # 14 项断言（ObjType / 几何 / 标签 / 并入 / 图例）
python scripts/doc_consistency.py   # 文档-代码一致性
```

完整用法（嵌 Word 三路线、审计、23 条踩坑清单）见 **SKILL.md** 与 `references/`。

## 目录

```
scripts/    visio_lib.py 绘图引擎 · theme.py 三套主题 · docx_embed.py EMF/栏宽
            word_insert.py 占位符插图 · audit_align.py 对齐审计(+白名单)
            docx_audit.py 交付自查 · selftest.py 回归 · doc_consistency.py 一致性
references/ cells-reference.md CellsU 速查 · pitfalls.md 23 条避坑 · docx-embed.md 嵌图配方
assets/     example_flowchart.py 最小示例
```

## 已知限制

1. 自由端连接线（图例线 / 刻度 / 总线引出）选中时仍显示包围盒 —— 加 `ObjType=2` 会被
   Visio 存盘时的 1-D 最小尺寸夹取改坏几何，本库取「图形正确优先」。
2. `merge_texts()` 是启发式；要精确控制就用显式 `label=`。
3. 依赖本机 Visio + pywin32 + 中文字体，**非离线可移植**。
4. COM 会话需自行收尾（推荐 `with VDraw(...)`）；忘关会残留 `VISIO.EXE`。

## License

MIT
