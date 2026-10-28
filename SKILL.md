---
name: visio-diagram
version: "1.0.0-beta"
description: 通过 COM 驱动本机 Visio 绘制矢量图，并嵌入 Word 文档。当用户要求画架构图 / 流程图 / 泳道图 / 模块关系图 / 数据流向图，或要求把图嵌入 docx，或提到 Visio、vsdx、EMF 矢量图、技术方案配图、总体设计方案插图时使用。适用于中文工程技术文档配图（总体设计方案、技术方案、试验大纲、研究报告），产出原生可编辑的 .vsdx 加矢量 .emf 与预览 .png，并自带文字溢出自检与 docx 版式核对。
agent_created: true
---

# Visio 矢量绘图与文档嵌图

## 用途

用 Python + pywin32 驱动本机 Visio，程序化生成**原生可编辑**的矢量图，并嵌入 Word 文档。
产出三件套：`.vsdx`（原生，可继续在 Visio 里改）· `.emf`（矢量，供 Word/PPT 嵌入）· `.png`（预览）。

适用于需要**精确控制版式**的中文工程文档配图。不适用于随手画示意图（那种场景用内联 SVG 更快）。

## 何时使用

触发场景：

- 用户要求画架构图、流程图、泳道图、模块关系图、数据流向图、时序链路图
- 用户要求把图嵌入 docx，或提到「配图」「插图」「图 N-x」
- 用户提到 Visio、vsdx、EMF、矢量图
- 需要为技术方案 / 总体设计方案 / 试验大纲 / 研究报告批量出图

不适用：

- 只需要一张快速示意（用内联 SVG/widget 更快）
- 需要照片级位图（走文生图工具）

---

## 前置检查（每次开工前先做）

```bash
# 1) Visio 是否安装
ls "C:/Program Files/Microsoft Office/root/Office16/VISIO.EXE"
# 2) 环境自检 —— ★ 一次查清：Python 版本 / pywin32 / Pillow 能否「真的取字体度量」
#    （★ 别只做 `import PIL`，它不加载 ImageFont 扩展，预览版 Python 下会假通过）
python -c "import sys; sys.path.insert(0,'scripts'); import visio_lib; ok,p=visio_lib.check_env(); sys.exit(0 if ok else 1)"
```

**★★ Python 版本要求（实测踩坑）**

| 版本 | 可用性 |
|---|---|
| **3.9 ~ 3.12** | ✅ 推荐。Pillow 11.x 稳定 |
| 3.13 | ✅ 可用（Pillow 11.x / 12.x 均可） |
| **3.15 alpha 等预览版** | ❌ **不要用**。Pillow 12.3 导入 `ImageFont` 会直接崩（`slot ID 85`），且**只在真正加载字体时才触发** |

★ **为什么前置检查必须用 `from PIL import ImageFont`**：旧版只做 `import PIL`，
  而 `PIL` 包的 `__init__` 并不加载 `ImageFont` 的 C 扩展 —— 于是预览版 Python 下
  这一步**能过**，等到 `_font()` 真正调 `ImageFont.truetype()` 才崩，排查成本很高。
  `visio_lib` 顶层已加同名自检（见 `check_env()`），失败时给出明确版本建议。

缺依赖则 `pip install pywin32 pillow`。**不想污染本机环境**时用 `uv` 起临时环境：

```bash
uv venv --python 3.12 .venv-vision && .venv-vision/Scripts/pip install pywin32 pillow
.venv-vision/Scripts/python scripts/selftest.py
```

**PowerShell 无法实例化 COM**（安全策略拦截）—— 必须走 Python。详见 `references/pitfalls.md` 第 1 条。

---

## 工作流

### 第 1 步：确定画幅

先量出目标文档的正文栏宽，**按 1:1 出图**（不要按 A4 满宽画，否则缩放后字号不可读）。

```bash
python scripts/docx_embed.py 目标文档.docx      # 输出「正文栏宽 17.50 cm」
```

没有目标文档时，默认按 **150~175 mm** 宽设计画布。

### 第 2 步：写绘图脚本

在项目下建 `绘图/` 目录，每图一个脚本。骨架照抄 `assets/example_flowchart.py`：

```python
import sys, os
sys.path.insert(0, "<skill>/scripts")
from visio_lib import VDraw

def main():
    # ★ with 退出（含异常）自动 close()，不会残留 VISIO.EXE；也可用 try/finally 手写
    with VDraw(150, 108, theme="tech-cn") as d:   # 画布 mm
        d.lane(0, 0, 150, 74, "处理链路")
        d.step_box(6, 12, 42, 26, "① 环节一", "说明文字")
        d.h_arrow(48, 25, 60)
        d.check()                                 # 文字溢出自检
        d.export_all("图表素材/图N1_某流程图")
```

坐标约定：**页面左上角为原点、Y 轴向下、单位 mm**。

可用的图元：

| 方法 | 用途 |
|---|---|
| `rect(x, y, w, h, **style)` | 矩形 |
| `text(x, y, w, h, s, size=, color=, align=)` | 文本框 |
| `line(x1, y1, x2, y2, arrow=4, lpat=1)` | 直线（`arrow=4` 实心箭头，`0` 无） |
| `poly([(x,y),...], arrow=4, lpat=4)` | 折线，仅末段带箭头 |
| `box(x, y, w, h, 标题, 说明)` | 标题 + 说明方块 |
| `step_box(x, y, w, h, 标题, 说明)` | 流程图用的环节框 |
| `lane(x, y, w, h, 标签)` | 泳道 / 分区（虚线框 + 左侧标签） |
| `h_arrow` / `v_arrow` | 水平 / 垂直箭头 |
| `legend(x, y, w, [(文字, lpat), ...])` | 图例行（条目在总宽 `w` 内**均布**；`w` 必须给足，太小或内容越出画布 → 导出按内容裁剪，右侧凭空多出一片空白，见 §22） |
| `add_cp(形状, fx, fy)` / `side_cps(形状, "top bottom")` | 给形状加连接点（**fy 自上边起算**），返回行号 |
| `connect(形状A, cpA, 形状B, cpB, via=[...])` | **单条折线连接线**，两端胶到连接点、按 via 折点走；挪框时端点跟随 |
| `connect(..., label=, label_at=, label_gap=, label_side=)` | 连接线**自带文本**。★ 标签一律用**沿线相对定位**（`label_at` 默认取最长段中点、`label_gap` 垂距、`label_side` 偏哪侧），**不要**手写 `label_pos` 绝对坐标（不随线移动，改布局就漂，见 §23） |

样式参数用**主题字段名**而不是写死颜色：`fill="muted_fill"`、`color="accent"`、`line="mgray"`。

内置主题（`theme=` 参数）：

| 主题 | 适用 |
|---|---|
| `tech-cn` | 中文工程文档：黑白灰线 + 深蓝标题 + 外部对象灰底（**默认**） |
| `mono` | 严格黑白印刷 |
| `blue` | 汇报材料 / 幻灯片 |

### 第 3 步：自检 → 修 → 出图

**Visio 不会因为文字超框而报错**，它会直接把字溢出到框外。必须自检：

```python
d.check()
# [文字溢出] 2 处
#   模板管理  8.5pt 框宽34.0 框高26.0 → 4 行需 31.2mm
# [连线自检] 出线方向 / 穿越实体框 均无
```

`d.check()` 现在同时查**连线**（2026-10-17 起）：

- `[连线方向] … 胶在形状#N 的 left 侧，首段却朝反方向` —— 连接线从框的**错误一侧**出线，
  会横穿整个框（框有底色时被压在下面，肉眼几乎看不出）。`connect()` 在写入时就会打印这条。
- `[连线穿越] 线#N 穿入 形状#M … 内部` —— 折线穿过了某个**有底色**实体框（泳道/边界框这类
  容器、无底色通栏文本框不算）。详见 `references/pitfalls.md` §24。

修法：**优先加高框体，其次减少文字**，最后才考虑缩字号（缩了在缩放嵌入后不可读）。
经验值：框标题 8.5~9 pt、框内说明 7~7.5 pt、图元左右内边距 0.8 mm。

出图后还要查**箭头有没有漂移**（箭头连到了框，但落点偏离框边中点，看起来就是歪的）：

```bash
python scripts/audit_align.py 图表素材/图N1.vsdx
```

它会报出「单独连到某条边、却偏离该边中点 > 2mm」的端点。这是**启发式筛查**：
若某框从同一侧扇出多条箭头（各自接到不同高度的框），偏移属正常，需结合渲染图目视判断。

**已定性为「刻意偏移」的端点**可登记进白名单（`scripts/audit_whitelist.json`），批量回归不再计失败：

```bash
python scripts/audit_align.py --add 图表素材/图N8….vsdx   # 把本次报出的偏移登记进白名单
python scripts/audit_align.py 图表素材/*.vsdx            # 此后已定性图 → 白名单豁免，exit 0
```

匹配键 = 图名（去扩展名）+ 端点坐标（±3mm 容差）+ 边名；命中打印「→ 白名单（登记时的 note）」；
条目在图中不再命中（布局改了 / 已修正）→ 结尾提示「已失效，建议清理」。
★ 登记 = 免检承诺，**先人工定性再 --add**，别拿它消音真实缺陷。

> 想直接看出偏差，可把 vsdx 的精确几何用红线叠到导出的 PNG 上对比 ——
> 红线与黑线重合说明几何无误，错开则说明坐标写错了（见 `references/pitfalls.md` §11）。

迭代到 `[文字溢出] 无` 且箭头对齐审计通过再定稿。详见 `references/pitfalls.md` 第 3、4、11 条。

### 第 3.5 步：文字并入宿主（默认自动）

「矩形 + 独立文本框」是常见写法，但**矩形自己就能写字**，连接线也自带 `Text`。
`VDraw` 默认在 `save()` / `export_all()` 前自动调用 `merge_texts()`，把
**落在矩形内的文本框**并入矩形、把**贴着连接线的小标签**并入连接线，**版式逐像素不变**
（12 图回归：逐瓦片中位位移 0.00px、p95 1px）。形状数可省 ~47%。

```python
d.merge_texts()                  # 通常不用手调，save()/export 会自动跑
d.merge_texts(verbose=True)      # 打印：[文字并入] 宿主 9 个 / 并入文本框 15 个 ｜ 连接线标签 3 个 …
```

开关与调试：

| 手段 | 作用 |
|---|---|
| `VISIO_MERGE=0` 环境变量 | 关闭自动并入 → 出「未合并」对照图做逐像素回归 |
| `VDraw.default_auto_merge = False` | 同上，代码里关（子类可覆写） |
| `VISIO_MERGE_DEBUG=1` 环境变量 | 打印每个文本框的去向与连接线标签的命中距离 |

规则：文本框**完全落在**某矩形内 → 并入（多个嵌套取最小者）；矩形已有文字则跳过；
连接线标签须**宽 ≤44mm 且高 ≤12mm**（避免误吞分区标题/图例/说明段），
且标签中心到折线距离 ≤6mm。**并入后 `Text` 变成富文本**（逐段字号/颜色/段前距复原），
改字请改源脚本重跑，不要在 Visio 里手改。

> 原理与踩坑（富文本 `CellsSRC(3,r,7)`、真块高 `_BH_EXTRA`、PIL 字号取整致折行少算、
> 连接线 `TxtLocPinX` 默认为 0、局部原点≠Begin 点）见 `references/pitfalls.md` §18。
>
> ★ **想让某条连线自带标签，别靠上面的启发式去猜 —— 显式传 `label=`**：
>
> ```python
> d.connect(a, cp["a"]["right"], b, cp["b"]["left"],
>           label="双 10GE", label_color="gray",
>           label_at=0.5, label_gap=2.2, label_side="up")   # 沿线相对定位，随线移动
> ```
>
> 定位一律用 `label_at` / `label_gap` / `label_side`（沿线相对定位）；
> **不要**手写 `label_pos` 绝对页坐标 —— 它不随线移动，一改布局标签就漂。
>
> 1-D 连接线自带文本**选中时不画大矩形框**、随线移动；启发式放宽反而会把多个文本框
> 并到同一条线上互相覆盖、丢文字。详见 `references/pitfalls.md` §20。

### 第 4 步：嵌入 Word

**推荐路线 A** —— 文档里留占位符，画完批量替换（Word 原生支持 EMF，矢量保真）：

```markdown
{{FIG:N2}}
```

```bash
python scripts/word_insert.py 方案.docx --placeholder "{{FIG:N2}}" --image 图N2.emf
```

默认宽度 = 正文栏宽 − 5 mm，高度按 EMF 宽高比自动算。

**路线 B** —— 从零合成 docx 时手工拼 OOXML，见 `references/docx-embed.md`（含三个必踩的坑）。

**路线 C** —— 甲方要求「双击进 Visio 改图」时，把已插好的图**原地换成 OLE 对象**：

```python
rng, w, h = sh.Range, sh.Width, sh.Height   # 先取 Range 与「宽 + 高」
sh.Delete()
new = rng.InlineShapes.AddOLEObject(FileName=abs_vsdx)
new.LockAspectRatio = 0                     # 必须先关，否则后写的一维会带动另一维
new.Width = w                               # 宽度手动还原
new.Height = h                              # ★ 高度也必须手动还原
```

> ★ **只写 `Width` 会毁掉图幅**：`AddOLEObject` 按 **OLE 对象（Visio）自身固有纵横比**定尺寸，
> Word 会据此自行重算高度，结果与 EMF 图幅不一致 —— 实测图被压扁 **最高 +20%**（方框明显变扁）。
> 高度必须取「替换前那张图」的 `sh.Height`（它是按 EMF 图幅算出来的唯一可信比例）显式写回。
> 详见 `references/pitfalls.md` §13。

要点：`InlineShapes` 按文档顺序与图列表一一对应；嵌的是 `.vsdx`（所以要保留 `.vsdx`）；
换完必须 `TablesOfContents.Update()` 重排页码；最后去掉 `settings.xml` 的 `updateFields`。
完整写法与校验见 `references/docx-embed.md` 路线 C。

### 第 5 步：交付前核对

```bash
python scripts/docx_audit.py 方案.docx
```

静态检查（无需 Word）：XML 良构、Content_Types 声明、`mc:Ignorable` 前缀一致性、`tblW`/`tcW` 属性名、加粗残留、正文栏宽。

Word 实测：页数、**最宽表格 vs 栏宽**、每张图尺寸、表格内外粗体数。

判定口径：最宽表格 ≤ 正文栏宽；表格内粗体 = 0；表格外粗体 = 标题段落数（标题样式自带加粗，正常）。

---

## 关键约束（务必遵守）

1. **按 1:1 出图** —— 画布宽 ≈ 正文栏宽，否则嵌入后字号不可读。
2. **必须自检** —— `d.check()` 是出图流程的固定一环，`TextBounds` 不存在，没有别的办法发现溢出。
   它同时输出**连线自检**（出线方向 / 穿越实体框）；两条都必须为「无」才算过。
3. **必须收尾（`d.close()`）** —— 否则 VISIO.EXE 残留后台进程，累积后 COM 调用会超时。
   推荐 `with VDraw(...) as d:`（退出含异常自动 close），或 `try/finally` 里 close；
   `export_all()` 现在**失败会抛异常**（不再静默吞掉），别忘了 finally 兜底。
4. **颜色走主题字段名** —— 不要写死 `"RGB(31,78,121)"`，便于整体换主题。
5. **算栏宽要含装订线** —— 忽略 `w:gutter` 会让表格/图溢出页面。
6. **连接线用 `connect()` 单条折线 + 胶连接点** —— 不要用 `poly()` 多条线段拼（挪框时不跟随、Visio 里无法整体编辑）；也不要用 `Page.Layout()` 重排（会搬动已排好的方块）。
7. **改完库先跑自测** —— `python scripts/selftest.py`（19 项回归断言，含离线 EMF 头校验）；文档改动跑 `python scripts/doc_consistency.py`（防结论漂移）。**图例总宽 `w` 要给足**，否则整行越出画布（§22）。

---

## 已知限制（用之前先知道）

1. **自由端连接线（有一端不胶水）选中时仍会显示包围盒** —— 图例线 / 刻度 / 箭头 / 总线引出
   这类线若设 `ObjType=2`，Visio 会在 `SaveAs` 后按 1-D 最小尺寸夹取**改坏几何**（刻度线被缩短、
   引出线被加出斜段）。本库取「**图形正确优先**」：这类线**不加 `ObjType`**，代价就是选中时仍画框。
   两端都胶水的连接线不受影响（正常不显框）。详见 `pitfalls.md` §21。
2. **文字并入（`merge_texts`）是启发式** —— 按「文本框是否落在矩形内 / 标签离线距离」判断，
   边界情形可能误吞或漏并。**要精确控制就用显式 `label=`**（写进连接线自带文本），别依赖启发式。
3. **依赖本机 Visio + pywin32，不是纯离线可移植** —— PowerShell 无法实例化 COM（沙箱策略），
   必须走 Python；换机器需装 Visio 与 `pywin32` / `pillow`，并确保中文字体（默认「微软雅黑」）。
4. **COM 单实例，进程需自行收尾** —— 忘 `d.close()` 会残留 `VISIO.EXE`，累积后 COM 调用超时；
   残留进程还会让 Word 嵌 OLE 时**复用旧文档**（`pitfalls.md` §14）。

---

## 资源

### scripts/

| 文件 | 用途 |
|---|---|
| `visio_lib.py` | 绘图引擎：`VDraw` 类（支持 `with`，退出自动 close）+ `text_extent_mm()` 字体度量 + `check_env()` 环境前置检查 |
| `theme.py` | 三套配色主题 + 字体文件映射 |
| `docx_embed.py` | `emf_size_mm()` 读 EMF 图幅 · `content_column_mm()` 算正文栏宽（**默认取末节、支持多栏折算**）· `ensure_content_type()` 补声明 |
| `word_insert.py` | 路线 A：Word COM 按占位符插图（含命令行） |
| `audit_align.py` | 箭头对齐审计：查出「箭头连到框但落点偏离边中点」的漂移；**已定性刻意偏移可用 `--add` 登记白名单**（`audit_whitelist.json`，坐标 ±3mm 匹配、失效提示） |
| `docx_audit.py` | 交付前自查：静态 OOXML 检查 + Word 版式实测 |
| `selftest.py` | **库自测**：画覆盖 A–L 12 类回归项的测试图 → 断言 ObjType / 几何点数 / 连接线标签 / 文字并入 / 图例不越界 / EMF 头重写（**改完库先跑它**，`python scripts/selftest.py`） |
| `doc_consistency.py` | **文档-代码一致性**：废止表述黑名单 + 关键结论必须在位 + 成对约束，防「代码改了、文档还写旧结论」的漂移 |

### references/

| 文件 | 内容 |
|---|---|
| `cells-reference.md` | Visio CellsU 属性速查、坐标换算公式、编号取值 |
| `pitfalls.md` | 29 条避坑清单（COM 沙箱、TextBounds 不存在、溢出、1:1、装订线、EMF 尺寸、COM 忙错误、箭头漂移、折点压邻居框、**OLE 图幅被压扁**、**残留 VISIO.EXE 致 OLE 嵌旧内容**、**连接点 Y 向上致 fy 颠倒**、**折线几何须平移进包围盒**、**单条折线连接线配方**、**文字并入宿主**、**上下相邻框连线端点选边**、**连接线自带文本 label=**、**选中折线出现矩形外框（根因 `Misc.ObjType` 未设为 2）**、**图例 `legend()` 忽略 w 致右侧一大片空白**、**连线标签离老远 —— 绝对坐标/并入不贴线**、**图变扁/比例不对 —— 先做三方对照（显示框·bounds·vsdx）**、**Word 存档把 72pt 规范化成 1in**、**静态 EMF 图文字被裁只能整张重绘**、**居中 / 目录页码这两个测量本身的坑**…） |
| `docx-embed.md` | 两条嵌图路线 + 三个 OOXML 坑的完整修复代码 |

### assets/

| 文件 | 内容 |
|---|---|
| `example_flowchart.py` | 最小可运行示例（约 60 行画一张带闭环与图例的流程图），照抄改内容即可 |

---

## 快速排查

| 现象 | 原因 | 位置 |
|---|---|---|
| COM 实例化被拒 | PowerShell 策略拦截 | `pitfalls.md` §1 |
| `TextBounds` AttributeError | 该 API 不存在 | `pitfalls.md` §2 |
| 导出后文字跑出框外 | Visio 不报错，需自检 | `pitfalls.md` §3 |
| 箭头看起来歪了 / 快贴到框角 | 落点偏离框边中点，用 `audit_align.py` 查 | `pitfalls.md` §11 |
| 嵌入后字太小 | 画布按 A4 满宽画了 | `pitfalls.md` §4 |
| 表格右侧溢出页面 | 漏算装订线 gutter | `pitfalls.md` §5 |
| Word 报「文件已损坏」 | Content_Types 或 mc:Ignorable | `docx-embed.md` 坑 1 / 坑 2 |
| 页数暴涨、每列一字 | `tblW`/`tcW` 写成 `w:val` | `docx-embed.md` 坑 3 |
| 嵌图后大小不对 | 采信了 EMF 的绝对尺寸（应只用宽高比） | `pitfalls.md` §9 |
| **换 OLE 后图被压扁** | **只写了 `Width`，Word 按 OLE 自身比例重算高度** | **`pitfalls.md` §13** |
| `Call was rejected by callee` | Word COM 忙，需重试 + 收尾不抛异常 | `pitfalls.md` §10 |
| VISIO.EXE 进程残留 | 忘了 `app.Quit()` | `pitfalls.md` §7 |
| **换 OLE 后某张图仍是旧内容** | **残留 VISIO.EXE，OLE 服务器复用了旧文档** | **`pitfalls.md` §14** |
| 竖线贯穿两个方块 / 连接点落在框的另一侧 | `Connections.Y` 是局部坐标 Y 向上，`fy` 必须自上边起算 | `pitfalls.md` §15 |
| 选中折线时出现空白矩形 / 选中框不在线上 | 几何行没平移进 `[0,Width]×[0,Height]`，须配 `LocPin` 拉回 pin | `pitfalls.md` §16 |
| **选中折线时沿包围盒出现一圈矩形外框** | **`Misc.ObjType` 没设成 `2`（可布线 1-D 形状）。须 `ConFixedCode=3` → `ObjType=2` → **回读局部原点**重写几何（顺序敏感）；短形状 / 自由端须回退** | **`pitfalls.md` §21** |
| **图整体变扁 / 长宽比不对** | **显示框照 EMF `frame` 算了（frame 的 X/Y DPI 不一致，约 1.914×）。先三方对照：显示框比 / bounds 比 / vsdx 并集比，再决定改谁** | **`pitfalls.md` §25** |
| 审计报「显示框 340x0」「长宽比 −100%」 | Word 保存会把 `72.0pt` 规范化成 `1in`，解析 style 必须支持 pt/in/mm | `pitfalls.md` §26 |
| **图内某框标题少一个字 / 文字压框（静态 EMF，无 vsdx 源）** | **EMF 里文字是矢量轮廓，PDF 提取不到；度量法阈值不可靠。只能重画整张** | **`pitfalls.md` §27** |
| 图不居中 / 图表目录页码对不上 | 测量方法本身有坑（`Range.Information(5)` 无效、页脚块不能按行序取） | `pitfalls.md` §28 |
| 竖线贯穿下方/上方那个方块 | 上下相邻两框连线须取「下框 `top` → 上框 `bottom`」 | `pitfalls.md` §19 |
| **选中连线标签时出现大矩形框** | 标签是 2-D 文本框，应写成连接线自带文本（`connect(..., label=)`） | **`pitfalls.md` §20.0** |
| 连线标签跳位 / 随线旋转 | 写了 `Control.TextPosition`（Visio 按页面坐标解释） | `pitfalls.md` §20.2 |
| **两字标签被挤成「一字一行」** | 连接线文本块有左右页边距，`TxtWidth` 不够 → 须清零页边距并留余量 | **`pitfalls.md` §20.3** |
| **某条连线标签整条文字消失** | 放宽 `_merge_line_labels` 启发式后多个文本框并入同一线、互相覆盖 | **`pitfalls.md` §20.4** |
| **连线标签离连接线老远**（改布局后尤其明显） | 写了 `label_pos` 绝对坐标（不随线移动），或标签是独立文本框、偏移被原样并进连接线。改用 `label_at/label_gap/label_side` 沿线相对定位；并入路径已自动贴线归一 | **`pitfalls.md` §23** |
| **图上右侧一大片空白 / 图例条目间距悬殊** | `legend()` 的 `w` 给得太小（或图例/长文本越出画布），导出按内容裁剪把越界部分也裁进来 | **`pitfalls.md` §22** |
| `/skill名` 命令载入的是**旧版**内容 / base directory 指向 `*.bak_*` | skills 目录里混入了 skill 的**备份副本**（会被扫描器当候选）。把备份移出 skills/（如 `~/.workbuddy-ai/skill-backups/`），只留正本 | 本条 |
