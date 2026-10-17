# 避坑清单

以下每一条都是实际踩过、且**不看就会浪费大量时间**的问题。按严重程度排序。

---

## 1. PowerShell 无法实例化 COM（环境级）

**症状**

```
[SandboxError] COM object instantiation can run arbitrary code ...
```

或 `New-Object -ComObject Visio.Application` 直接失败。

**原因**：执行策略把 COM 实例化当作任意代码执行拦截了。

**对策**：**一律改用 Python + pywin32**。

```bash
pip install pywin32
```

```python
import win32com.client as win32
app = win32.Dispatch("Visio.Application")
```

用 late binding（`Dispatch`）即可，不必 `gencache.EnsureDispatch`。

**推论**：先探测环境，不要假设「本机没装 Visio」。检查：

```bash
ls "C:/Program Files/Microsoft Office/root/Office16/VISIO.EXE"
```

Office 通常同时装 Visio 与 Word，Word 可用于更新目录域、转 PDF、做版式实测。

---

## 2. `Shape.TextBounds` 不存在（API 级，最隐蔽）

**症状**：想判断文字有没有溢出框，调用 `shape.TextBounds`：

```
AttributeError: <win32com.gen_py...> object has no attribute 'TextBounds'
```

早晚绑定都取不到。**这个属性/方法在 Visio 类型库里根本不存在**。

**危险之处**：如果把它包在 `try/except` 里「探测」，会**静默失败** —— 检查器永远返回「无溢出」，让人误以为排版没问题。

**对策**：改用**字体度量**自行估算折行高度。

```python
from PIL import ImageFont
f = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", px)   # 字号 -> 像素
wpx = width_mm / 25.4 * 96.0
# 按框宽逐字试排，累计行数
asc, desc = f.getmetrics()
need_mm = lines * (asc + desc) / 96.0 * 25.4
```

已封装为 `visio_lib.text_extent_mm()` 与 `VDraw.check()`。

**注意**：字体名要和度量用的字体文件对应（微软雅黑 → `msyh.ttc`，宋体 → `simsun.ttc`），否则行数估算会偏。

---

## 3. 文字溢出只能靠自检发现

Visio **不会**因为文字超框而报错或警告 —— 它会直接把文字溢出到框外，导出后才发现。

**对策**：每张图导出前跑一次自检，把 `check()` 作为出图流程的固定一环。

```python
d.check()     # 打印 [文字溢出] N 处，含「几行 / 需要多高」
```

**调参经验**（本 skill 的实际取值）：

- 图元内左右内边距 `MARGIN_MM` 从 1.2 → **0.8**，可显著减少假性溢出；
- 中文字号：框标题 8.5~9 pt，框内说明 7~7.5 pt；再小在 A4 正文里就糊了；
- 宁可**把框加高**，不要缩字号 —— 缩字号在缩放嵌入后会不可读。

---

## 4. 按 1:1 出图，不要按「画布大小」出图

如果按 A4 满宽画（比如 21 cm），嵌进正文栏宽只有 17.5 cm 的文档时会被缩到 83%，字号 7 pt 实际只有 5.8 pt，**印出来看不清**。

**对策**：**按最终印刷宽度 1:1 出图**。

- 先量出正文栏宽（见 `docx_embed.content_column_mm()`）；
- 图幅按栏宽的 95%~100% 设计（例如栏宽 175 mm → 画布宽 150~175 mm）；
- 嵌入时按 100%~105% 缩放，字号的视觉大小基本不变。

---

## 5. 忽略装订线（gutter）导致表格/图溢出页面

**症状**：文档页宽 21 cm，左右边距各 1 cm，按 19 cm 排表格 —— 渲染 PDF 后文字跑到 21.19 cm，**右侧溢出页面**。

**原因**：模板的 `w:pgMar` 里带 `w:gutter="851"`（1.5 cm 装订线），**Word 中是生效的**。正文实际起点 = 左边距 + 装订线 = 2.5 cm，可用栏宽 = 21 − 1 − 1 − 1.5 = **17.5 cm**。

**对策**：从 `sectPr` 动态算，不要硬编码。

```python
from docx_embed import content_column_mm
mm, twips = content_column_mm("方案.docx")     # -> (175.0, 9921)
```

注意：若带 `w:gutterAtTop`，装订线在上边，不减宽度。

---

## 6. 中文行末标点悬挂会被误判成「溢出」

渲染 PDF 后用文字块边界量宽度，会看到行末超出栏宽约 **0.2 cm**。**这不是缺陷** —— 是 Word 的「允许标点溢出边界」中文排版行为（`、。；，）` 等标点可以悬挂在边界外）。

**对策**：判定溢出时给 0.3 cm 容差；或只统计**非标点结尾**的行。

---

## 7. Visio 进程残留

`doc.Close()` 之后如果忘了 `app.Quit()`，`VISIO.EXE` 会留在后台。反复运行脚本会累积多个实例，最终导致 COM 调用超时或失败。

**对策**：`close()` 里同时做 `Close()` 和 `Quit()`，并且都包 `try/except`（文档可能已被关闭）。

```python
for fn in (lambda: self.doc.Close(), lambda: self.app.Quit()):
    try:
        fn()
    except Exception:
        pass
```

另外设 `app.AlertResponse = 7` 抑制弹窗 —— 否则一个「是否保存」对话框就能让无人值守的脚本永久挂起。

---

## 8. `Char.Font` 不加引号会静默失效

```python
s.CellsU("Char.Font").FormulaU = '"微软雅黑"'   # ✓
s.CellsU("Char.Font").FormulaU = '微软雅黑'      # ✗ 不报错，但字体不生效
```

Visio 公式里字符串要带引号。同理 `CellsU` 的值一律是字符串，布尔用 `"0"`/`"1"`。

---

## 9. EMF 的绝对物理尺寸不可靠，只能信宽高比

**症状**：从 EMF 头 `rclFrame` 读出 115.13 mm，但同一张图导出的 PNG 是 530 px @ 96 DPI = **140 mm**。差了 22%，到底哪个对？

**原因**：Visio 导出 EMF 时写的 `rclFrame` 隐含约 **117 DPI**，而 PNG 用 96 DPI。EMF 的 `rclFrame`（0.01 mm）与 `rclBounds`（像素）之间并不遵循 96 DPI 约定，不同导出器取值也不同。

**验证方法**：同时读 `rclBounds` 和 `rclFrame`，两者推出的隐含 DPI 若不是 96，就说明绝对值不可信。

```python
info = emf_info("图.emf")
info["implied_dpi"]     # 117.4 → 绝对值不可信
info["ratio"]           # 1.0843 → 这个可靠
```

**对策**：**只用宽高比**，宽度按目标栏宽显式指定，高度 = 宽度 × ratio。

```python
w_mm = 170.0
h_mm = w_mm * info["ratio"]
```

实测四张图的比对结果 —— 比例吻合到小数点后三位，绝对值差 22%：

| 图 | EMF `frame_mm` | PNG 宽高比 | EMF 宽高比 |
|---|---|---|---|
| N2 | 113.4 × 116.8 | 1.0306 | 1.0302 |
| N3 | 115.1 × 124.8 | 1.0849 | 1.0843 |
| N5 | 118.4 × 118.4 | 1.0000 | 0.9998 |
| N6 | 115.6 × 116.8 | 1.0132 | 1.0110 |

---

## 10. Word COM 会抛「Call was rejected by callee」

**症状**：

```
pywintypes.com_error: (-2147418111, 'Call was rejected by callee.', None, None)
```

出现在 `Documents.Add()` 之后紧接着的赋值、`Close()`、`SaveAs()` 等看似普通的调用上。

**原因**：Word 正在忙（后台加载、重排、或弹了一个**模态对话框**）。此时 COM 调用会被直接拒绝。

**危险之处**：如果这个异常发生在 `finally` 块里，它会**掩盖主流程的真实结果**；而且 `Quit()` 没执行到，`WINWORD.EXE` 会残留在后台 —— 残留实例又会继续抢占，导致后续每次调用都失败，形成恶性循环。

**对策**：两件事一起做。

```python
def _retry(fn, times=6, delay=0.4):
    for i in range(times):
        try:
            return fn()
        except Exception:
            if i == times - 1:
                raise
            time.sleep(delay)

def _quiet(fn):
    try:
        fn()
    except Exception:
        pass

doc = _retry(lambda: app.Documents.Open(path))    # 关键调用加重试
...
finally:
    _quiet(lambda: doc.Close(False))              # 收尾永不抛
    _quiet(lambda: app.Quit())
```

**排查残留进程**（Windows）：

```bash
tasklist //FI "IMAGENAME eq WINWORD.EXE"
taskkill //PID <pid> //F
```

Visio 侧同理，但 `visio_lib.VDraw.close()` 已做 `_quiet` 处理，实测不会残留。

**注意**：`Documents.Add()` 后立刻操作 `Content.Text` 是最容易触发的位置 —— 新建文档时 Word 还在初始化。

---

## 11. 箭头「漂移」——连上了框，但落点不在边中点

**症状**：箭头确实连到了目标框，但落点明显偏向一侧（常见是快贴到框角），看起来像箭头「漂移」了。

**真实案例**：一张跨泳道图，源框 `142–164`（中心 153）、目标框 `100–152`（中心 126）。
作者为了取巧，把竖箭头放在 `x=147` —— 那是**两个框水平重叠区（142–152）的中点**，
看上去「两边都不偏」。但实际渲染出来：

- 在源框上只占 **23%** 位置（贴着左边）
- 在目标框上占到 **90%** 位置（快贴到右沿）

两端同时偏离框中心，视觉上就是「箭头向右漂移」。

**为什么容易犯**：取重叠区中点这个思路本身有道理（保证箭头两端都在框内），
但它忽略了「视觉上箭头应对准框中心」这个更强的约束。

**对策**：两个框错位时，**不要用直连**，改用 L 形折线，两端各对准框中心。

```python
# ✗ 直连：两端都偏
d.v_arrow(147, 38.5, 81.5)

# ✓ L 形：源框底边中心 → 下 → 左 → 目标框顶边中心
d.poly([(153, 38.5), (153, 77), (126, 77), (126, 81.5)], arrow=4)
```

**自检**：`scripts/audit_align.py` 会报出「单独连到某条边、却偏离该边中点 > 2mm」的端点。
注意它是**启发式**：同一侧扇出多条箭头（各自接到不同高度的框）时，偏移是正常的。

**排查渲染 vs 设计**：若怀疑是渲染问题而非坐标问题，把 vsdx 的精确几何用红线叠到导出的 PNG 上：

```python
# 读 vsdx 的 PinX/PinY/Width/Height 与 Begin/End，按同一比例画到 PNG 上
dr.rectangle([x1*kx, y1*ky, x2*kx, y2*ky], outline=(255, 0, 0), width=1)
```

红线与黑线重合 → 几何无误，是设计坐标的问题；错开 → 是坐标换算或渲染的问题。

**顺带一提**：箭头端点与框边之间常留 0.5mm 左右间隙（避免箭头压在框线上），
这是正常做法；但如果间隙大到看起来「箭头没顶到框」，就应缩到 0.3mm 以内。

---

## 12. L 形折线的折点落在相邻框的边上

**症状**：`audit_align.py` 报某个带箭头线段的**起端**偏离某框边中点很远（几十 mm），
但你分明是按「对准框中心」写的。看渲染图会发现：折线有一段**贴着旁边那个框的下沿**走。

**真实案例**：四个结果框并排，`未命中`(x 122–166) 要连到 `按配置未输出`(x 94–134)。
写成

```python
# ✗ 折点 (114, 60.5) 正好落在邻居框「命中·数值差>容差」(x 64–116, 下沿 y=60) 的边上
d.poly([(138, 60.5), (114, 60.5), (114, 69.5)], arrow=4)
```

`x=114` 落在左邻框的 x 区间 `64–116` 内，而 `y=60.5` 几乎就是它的下沿 `y=60`。
于是：折线的水平段贴着左邻框下沿画过去（视觉上像「粘」在别人框上），
同时审计把起端 `(114, 60.5)` 判定成「连到左邻框下边、却偏离中点 24mm」。

**根因**：只想着「折点要落在两个框之间的空隙里」，忘了检查**折点是否落在别的框的边上**。
横向折线的 `y` 选得离某排框的下沿太近（差 0.5mm）时尤其容易中招。

**对策**：折点所在的那条「走线车道」要和上下两排框都留出**明显间隙**（≥3mm），
并且起端/末端各对准源框与目标框的边中点：

```python
# ✓ 走 y=65 车道：距上排框下沿(60) 5mm、距下排框上沿(70) 5mm；两端各对框中心
d.poly([(144, 60.5), (144, 65), (114, 65), (114, 69.5)], arrow=4)
```

**自检**：`audit_align.py` 会把这类问题暴露成「起端偏离很远」——
**偏离特别大（>20mm）时，先怀疑折点压到了邻居框的边**，而不是怀疑坐标换算。
同一批 L 形折线尽量**共用同一条走线车道**（同一 `y`），既整齐又便于一次性检查。

---

## 13. ★ 换成 OLE 后图幅被压扁（只写 `Width` 的陷阱）

**症状**：EMF 直接插入时比例正常；一旦「原地换成 Visio OLE 对象」（路线 C），
图就**变扁**了 —— 方框被横向拉长 / 纵向压扁，肉眼一眼可见。实测 11 张图里 6 张失真，
最严重 **+20%**（例如 N4：EMF 宽高比 0.973，文档里变成 1.172）。

**根因**：`InlineShapes.AddOLEObject()` 插入时按 **OLE 对象（Visio 文档）自身的固有纵横比**
决定尺寸。脚本若只写 `new.Width = w`，Word 会**自行重算高度** —— 算出来的是 Visio 那边的比例，
不是 EMF 的比例。于是「宽对了、高错了」。

**为什么难发现**：页数、表格数、禁词、`w:object` 计数、`Content_Types`、孤行标题
**全都正常**；`docx_audit.py` 也不查这个。必须专门比对图幅才能发现。

**对策**：替换前记下原图的 `Width` **和** `Height`，替换后**两个都写回**，并且先关掉纵横比锁定
（否则后写的那一维会带动另一维）：

```python
rng, w, h = sh.Range, sh.Width, sh.Height     # ← 原图幅来自 EMF 宽高比，唯一可信
sh.Delete()
new = rng.InlineShapes.AddOLEObject(FileName=abs_vsdx)
new.LockAspectRatio = 0                       # msoFalse，必须先关
new.Width = w
new.Height = h                                # ★ 关键：显式锁死高度
```

**校验（两种口径，任一即可）**：

```python
# 口径 1：查 XML 里 v:shape 的 style，与 EMF 宽高比比对，偏差应 ≤0.1%
#   <v:shape id="_x0000_i1026" type="#_x0000_t75" style="width:481.5pt;height:495.5pt">
re.findall(r'style="width:([0-9.]+)pt;height:([0-9.]+)pt"', document_xml)

# 口径 2：导出 PDF 后量绘图矩形包围盒（pymupdf）
rs = [d['rect'] for d in page.get_drawings()]      # 取图所在 y 区间的那些
aspect = (max(r.x1 for r in rs) - min(r.x0 for r in rs)) / \
         (max(r.y1 for r in rs) - min(r.y0 for r in rs))
```

> ⚠️ 量 PDF 时**别用固定 clip 框**：图比 clip 高时会把下半截切掉，算出来的宽高比偏大，
> 让人误以为还没修好。要么用 `get_drawings()` 的矩形包围盒，要么先确认 clip 覆盖整张图。

**顺带**：`docx_embed.emf_size_mm()` 读出的 `rclFrame` 宽高比是可靠的（见 §9），
`合成`/`插入` 阶段按它算出的宽高是**对的**；错只错在「换 OLE」这一步。

---

## 14. 残留 `VISIO.EXE` 会让 OLE 嵌入**陈旧内容**（2026-10-10 实坑）

**现象**：重跑图脚本 + 重跑合成/定稿后，docx 里某张图**仍是旧内容**（旧文字/旧版式），
而 `图表素材/<图>.vsdx` 磁盘文件**明明是新画的**；「合成后、未定稿」的内嵌 EMF 也是新的
→ 说明问题只出在**「换 OLE」这一步**。

**根因**：上次会话遗留了一个 headless `VISIO.EXE` 进程。Word 执行
`InlineShapes.AddOLEObject(FileName=vsdx)` 时，**Visio 的 OLE 服务器复用了那个已存在的实例**，
而该实例里还开着**旧版文档**（同路径）→ 于是把旧内容嵌进了 docx。

**修复**：
```bash
tasklist //FI "IMAGENAME eq VISIO.EXE"          # 先看有没有残留
taskkill //PID <pid> //F                        # 杀掉（Word 也要一起清）
# 然后重跑：合成 → 定稿
```

**预防 / 复验**（嵌入后逐张核对，别只看「11 张」够不够）：
```python
import zipfile, io
z = zipfile.ZipFile(docx)
# ① OLE 预览：word/media/*.emf（矢量，文字是 UTF-16LE）
for n in z.namelist():
    if n.startswith('word/media/'):
        t = z.read(n).decode('utf-16-le', 'ignore')
        assert '旧词' not in t, n            # 有旧词 = 预览是旧的
# ② 内嵌源文件：word/embeddings/*.vsdx 本身就是 zip，读它的页面 XML
for n in z.namelist():
    if n.startswith('word/embeddings/'):
        iz = zipfile.ZipFile(io.BytesIO(z.read(n)))
        t = ''.join(iz.read(p).decode('utf-8', 'ignore')
                    for p in iz.namelist()
                    if p.startswith('visio/pages/page') and p.endswith('.xml'))
        assert '旧词' not in t, n            # 有旧词 = 嵌进去的 vsdx 是旧的
```

> 一句话：**动 Word/Visio COM 前，先确认没有残留的 `VISIO.EXE` / `WINWORD.EXE`。

---

## 15. ★ 连接点细胞的 Y 是「向上」的 —— `add_cp` 的 fy 上下颠倒（2026-10-12 实坑）

**症状**：竖线从「上框顶边」起笔，却**贯穿两个方块**才停下；COM 逐条打印
`BeginX/BeginY/EndX/EndY`（`Result("mm")`）与设计坐标完全不符。一次影响 12 张图里的
49 条连接线，肉眼却只在渲染图上看出「线穿过了框」。

**根因**：`Connections.Y{n}` 是**局部坐标**——原点在形状**左下角**、**Y 轴向上**。
绘图库对外约定是「左上角原点、Y 向下」，写 `=Height*fy` 时 `fy=1.0` 落在**上边**
（与直觉相反）。这条公式在类型库里看不出任何异常，Visio 也不报错。

**对策**：库里内部改写为 `=Height*(1-fy)`，`fy` 的语义改成「**自上边起算**」
（0 = 上、1 = 下、0.5 = 中）；`side_cps` 的方向向量同步（top → DirY=+1）。
调用方**零改动**。

```python
self._f(s, "Connections.X%d" % n, "=Width*%.10g" % float(fx))      # fx 自左起算，方向天然一致
self._f(s, "Connections.Y%d" % n, "=Height*(1-%.10g)" % float(fy))  # fy 自上起算，必须翻转
```

**自检口径**（比看图可靠）：连接线画完后用 COM 读每条线的
`BeginX/BeginY/EndX/EndY` 的 `Result("mm")`，逐条与设计坐标核对。
页高要从 `PageSheet.CellsU("PageHeight").Result("mm")` 读，**不要拍脑袋填**。

---

## 16. ★ 折线几何必须平移进 `[0,Width]×[0,Height]` —— 否则选中框脱离折线（2026-10-12 实坑）

**症状**：在 Visio 里选中一条折线连接线，**选中框（外接框）不在折线上**，
而是在别处出现一块透明空白矩形；拖动时图形行为也难以预测。

**根因**：几何行 `Geometry1.X{n}/Y{n}` 同样活在「原点左下、Y 向上」的局部坐标里。
折线向下/向左走时偏移为负；若直接把负偏移写进几何行，而 `Width/Height` 仍按
`max(xs)-min(xs)` 这种「跨度」写，**形状的宽度/高度就与几何的实际范围分离** ——
Visio 的选中框是跟着 Width/Height 算的，于是跑到 pin 的另一侧去了。

**对策**：把几何**整体平移到左下角**（页面几何位置与胶水端点完全不变），
并按 **Visio 标准 1-D 变换**设置 `Pin`/`LocPin` —— 即 `Pin` 与 `LocPin` **都落在包围盒中心**：

```python
xs = [p[0] for p in loc]; ys = [p[1] for p in loc]
dx0, dy0 = min(xs), min(ys)                    # 几何左下角
w, h = max(xs) - dx0, max(ys) - dy0
self._f(conn, "LocPinX", "%.10g mm" % (w / 2.0))            # 半宽
self._f(conn, "LocPinY", "%.10g mm" % (h / 2.0))            # 半高
self._f(conn, "PinX", "BeginX + %.10g mm" % (dx0 + w / 2.0))  # 包围盒中心（以 Begin 为锚）
self._f(conn, "PinY", "BeginY + %.10g mm" % (dy0 + h / 2.0))
self._f(conn, "Width",  "%g mm" % w)
self._f(conn, "Height", "%g mm" % h)
# 几何行写 (lx - dx0, ly - dy0)
```

★★ **不要**写成 `PinX = "=BeginX"` + `LocPinX = -dx0`（LocPin 落在包围盒**角**上）——
那会让形状的外接框与几何范围分离。**但注意：矩形外框的真因不是它**，
而是 `Misc.ObjType` 没设成 `2`（2026-10-13 单变量 A/B 推翻旧结论）。
详见 **§21**（2026-10-13 实坑）。

**XML 级验收**（不依赖 Visio GUI，vsdx 本身就是 zip）：

```python
# ElementTree 读 visio/pages/page*.xml；★ 注意：XML 的 <Row IX="n"> 从 1 开始，
# 且 CellsSRC 的段索引不含 component 行——用 CellsSRC 读会整体错位一行，以 XML 为准。
# <Cell N="X1" V="0.31496"/> 的 V 是英寸，×25.4 才是 mm。
```

判据（容差 0.02 mm）：① 所有几何行落在 `[0,Width]×[0,Height]`；② 外框四边与几何四边重合；
③ `LocPinX/Y` 落在框内；④ 首末折点换算回页面后 = `BeginX/Y` 与 `EndX/Y`；⑤ 无斜线段（全正交时）。

**连带好处**：外接框贴合后，`Page.Export` 的裁剪边界也变小了（不再被错误外框撑大）——
重跑图脚本后 PNG/EMF 的字节数变化属**预期**，不是内容改动。

---

## 17. 单条折线连接线配方（替代「多条组合线 + 不连点」）

Visio 里用多条独立 line 拼一条连接线，端点不连 connection point，后续挪框时线不会跟着走。
改成**单条折线连接线 + 胶到连接点**：

```python
# 1) 造 1-D 形状并胶到两端的连接点
conn = d.connect(src_shape, src_cp, dst_shape, dst_cp)      # 库已封装
# 2) 内部覆写坐标框架，把 pin 钉在 Begin 点
#    PinX="=BeginX"  PinY="=BeginY"  Angle=0  FlipX=FlipY=0
# 3) 几何行：段索引 10；第 1 行 MoveTo，其余 LineTo
conn.AddRow(10, -2, 139)        # 139 = visTagLineTo；RowCount(10) 含 component 行
```

**两条死路（实测否决）**：`Page.Layout()` 会搬动已排好的方块（`reroute()` 置空即可，
不要整体重排）；`pg.Drop(ConnectorToolDataObject)` 动态连接线规程下再 `AddRow` 会挂死。**

---

## 18. ★★ 文字并入宿主形状（`merge_texts`）：把「矩形 + 独立文本框」合并成「矩形自带文本」（2026-10-12）

动机：图里大量「矩形 + 一个文本框」的组合，框本身就能写字；连接线同理自带 `Text`。
合并后形状数可省 ~47%（实测 12 图：584 形状、302 纯文本框，其中 274 在矩形内）。

`VDraw` 默认 `default_auto_merge = True`，在 `save()`/`export()` 前自动调用
`merge_texts()`；用环境变量 `VISIO_MERGE=0` 或类属性 `VDraw.default_auto_merge = False`
可关闭（**用于出「未合并」对照图，逐像素回归**）。`VISIO_MERGE_DEBUG=1` 打印
每个文本框的去向（并入哪个宿主 / 无宿主 / 线标签命中与距离）。

### 18.1 富文本写入：**不要用 `Characters.CharProps`**

`Shape.Characters.CharProps(prop, newValue)` 的 `newValue` 在类型库里声明为**整数** →
**8.5pt 被截成 8pt、颜色索引完全写不进**（早/晚绑定都一样）。正解是**直接写 Character 段单元格**：

```python
# 段索引 3 = Character；列索引：0=Font / 1=Color / 7=Size
host.CellsSRC(3, row, 7).FormulaU = "8.5 pt"     # 字号
host.CellsSRC(3, row, 1).FormulaU = "RGB(89,89,89)"  # 颜色
```

- **行号 ↔ 段落映射（哨兵法）**：`CharProps(7, 41 + 3*i)` 给第 i 段写唯一哨兵字号，
  再遍历 `CellsSRC(3, r, 7).FormulaU` 反解出行号。**不要**假设「段 i = 行 i」。
- **段前距**：`Characters.ParaProps(4, pt)`（★ 方法名是 `ParaProps`，**不是** `ParagraphProps`）。
  作用范围是**传入范围内的所有段落** → 必须 `c.Begin = offs[i]; c.End = offs[i] + 1`
  **只取该段首字符**，否则会把后面所有段落的段前距一起改掉。

### 18.2 版式复原常数（实测标定，微软雅黑）

```
行高      LH = 0.42911·字号 − 0.0514  (mm)   ≈ 1.216×字号，残差 ≤0.1mm
首行墨迹   文字块顶 + 上边距 + 1.02 mm        (_INK)
真块高     n×LH + _BH_EXTRA，_BH_EXTRA = 0.52 mm
段间推进   A ≈ LH + _ADV_BONUS，_ADV_BONUS = 0.15 mm
```

- ★★ **真块高 ≠ n×行高**：Visio 做 `VerticalAlign=1/2`（居中）时用的「文字块高」比行高累加多
  **0.52 mm**（8.5pt 与 9.0pt 实测同值）。不补 → 整块**下移 0.26 mm**（半量）。
  `VerticalAlign=0`（顶对齐）不用块高 → 不受影响，故**只补非 0 对齐**。
- `_INK = 1.02` 在「块顶→首行墨迹」与「段间推进反解」中**同时出现 → 完全抵消**，
  所以它取什么值都不影响最终结果（但两个公式里必须一致）。
- 段间推进被 Visio **量化**（约 0.1pt ≈ 0.035mm）→ `_ADV_BONUS` 在 0.10~0.30 之间
  几乎无差别（残差已到量化底），取 0.15 即可。
- **实测证伪**：「块高超出可用高度时 Visio 不再居中而贴顶」——**不存在这个夹紧规则**
  （n8 反例）。别加这个补偿。

### 18.3 ★★ 根因陷阱：PIL 字体度量会把字号**取整** → 折行行数少算

`ImageFont.truetype(path, size)` 的 `size` 是**整数像素**；8.5pt 在 96dpi 下是 11.333px
→ **被取整成 11px**，`getlength()` 随之系统性偏小 **~3%**（Pillow 12.3.0 传 float 也照样取整）。

后果：`text_extent_mm()` 把「读时刻（时间戳计数寄存器）」13 字算成 1 行（PIL 143px <
限度 145.13px），而 Visio 在 38.4mm 宽里实际折 **2 行** → 块高算小 4.1mm → 居中后文字下移 1.8mm。

**对策**：按「真实 px / 取整 px」把长度放大回真实值：

```python
def _font_scale(size_pt, font_name="微软雅黑"):
    px = max(6, int(round(size_pt * 96.0 / 72.0)))
    return _font(size_pt, font_name), (size_pt * 96.0 / 72.0) / float(px)
# 之后 f.getlength(...) * sc 才是真实宽度
```

### 18.4 ★★ 连接线标签并入：局部原点**不是 Begin 点**

给 1-D 连接线写 `conn.Text = "下发"` 可行（文字出现在连接线中点），但要把文字**钉回原位**：

- **`TxtPinX/TxtPinY` 用形状局部坐标**（Y 向上）。而 §17 的 `_author_path()` 会把几何
  平移到左下角并用 `LocPin` 把 pin 拉回 Begin 点 → **局部原点 = (PinX−LocPinX, PinY−LocPinY)**，
  **不是 Begin 点**。写 `cx − BeginX` 只在 `LocPinX=0`（向右/向下走的线）时对。
  → 实测 图N9「上报」返程线 `Begin=(94,85.5)`、`LocPinX=18`，用 `Begin` 换算得
  距标签 **9.86mm**（超 tol 漏并），用局部原点换算才是真实的 **5.00mm**。
- **`TxtLocPinX` 默认 = `TxtWidth*0`**（文字块**左缘**贴文本针）→ 必须显式写
  `TxtLocPinX = w/2`、`TxtLocPinY = h/2`，否则文字**右偏 ~2mm**。
- `TxtAngle = 0` 保证文字始终水平（不随连线角度旋转）。
- 命中判据：标签中心到折线的**最短距离 ≤ tol（6mm）** + **尺寸守卫**（宽 ≤44mm、高 ≤12mm）
  → 只并「标签」，不误吞分区标题/图例/说明段。
- 匹配前先读全部 1-D 形状的几何折线；**读行数必须用 `s.RowCount(10)`**：
  - `s.RowCount("Geometry1")`（字符串）→ `invalid literal for int()`；
  - 对不存在的行 `CellsU("Geometry1.X3").Result("mm")` **返回 0.0 而不抛错** →
    「读到 None 就停」会读出一长串 (0,0) 假点（实测 180 个）。必须按行数截断。

### 18.5 回归方法：`VISIO_MERGE=0/1` 逐像素对照

同一脚本跑两遍（关闭/开启合并），PNG 逐像素比。**判据看「差异率」与「逐瓦片位移」**，
不要只看整图差异率（抗锯齿会抬高底噪）：

```python
# 1) 差异率：|a-b| > 60 的像素占比；12 图实测 0.27%~1.57%（底噪来自抗锯齿）
# 2) 逐瓦片垂直互相关位移：墨迹量 ≥60 的瓦片统计 → 12 图中位 0.00px、p95 1px、最大 2px
```

判定「版式没动」的**硬指标是逐瓦片位移**，不是差异率。残余差异来自抗锯齿与
段前距量化（0.035mm 级），视觉不可见。

---

## 19. ★ 上下相邻两框连线：端点必须取「下框 top → 上框 bottom」（2026-10-13 实坑）

### 19.1 症状

图 9-1 中「授时源」（下框）到「NTP 授时设备」（上框）的连接线**贯穿了整个 NTP 框**，
箭头停在 NTP 框的**外上方**；而预期只是两框之间的一段短连线。

### 19.2 根因

```python
cp["ntp"] = {"bottom": d.add_cp(b_ntp, 0.5, 1.0, diry=1),    # 下边 y=73（已定义）
             "top":    d.add_cp(b_ntp, 0.5, 0.0, diry=-1)}   # 上边 y=56
cp["ref"] = {"top": d.add_cp(b_ref, 0.5, 0.0, diry=-1)}       # 授时源上边 y=77
...
d.connect(b_ref, cp["ref"]["top"], b_ntp, cp["ntp"]["top"])    # ← 错：终点用了上框的 top
```

授时源在 NTP **下方**（授时源 y 77~94，NTP y 56~73）。终点写 `b_ntp.top`（y=56）后，
线段从 y=77 直画到 y=56 —— 中间 y∈[56,73] 恰好是 **NTP 框的内部**，于是「穿框」。
`cp["ntp"]["bottom"]`（y=73）当时已定义却**从未使用**，正是该用的那个。

### 19.3 修法

```python
d.connect(b_ref, cp["ref"]["top"], b_ntp, cp["ntp"]["bottom"])  # 下框顶边 → 上框底边
```

同时删掉已无用的 `cp["ntp"]["top"]` 定义（避免以后误用）。

### 19.4 判据：选端点前先问两句话

1. **两个框谁在上、谁在下？** 读 `step_box` 的 `ytop/h` 确认（本库 y 向下、左上原点）。
2. **线从哪个边出、进哪个边？** 上下相邻 → 出「下框 **top**」、进「上框 **bottom**」；
   左右相邻 → 出「左框 **right**」、进「右框 **left**」。

★ `add_cp(s, fx, fy)` 的 **fy=0 是上边 / fy=1 是下边**（见 §15）；`"top"/"bottom"` 只是
别名。**用 `side_cps()` 批量加四向连接点可彻底避免手写 fy 出错**：

```python
cp = d.side_cps(b_ntp, "top bottom")   # → cp["top"] / cp["bottom"]
```

### 19.5 验证三件套

- 重跑图脚本 → PNG 尺寸与改前**一致**（755×536，内容包围盒未变）、`[文字溢出] 无`
- 重跑管线 → **图幅偏差 `1.408 vs 1.409 / 0.096%` 与改前逐字一致** → 嵌入图框尺寸不受影响
- **从最终 PDF 渲染该页目视复核**（Word 会把 OLE 渲染进 PDF）—— 这是最可靠的确认方式

---

## 20. ★★ 连接线**自带文本**：`d.connect(..., label=...)`（2026-10-13 实坑 ×3）

### 20.0 为什么要这么做

「矩形 + 独立文本框」做连线标签时，**选中标签会画出一个大矩形框**（2-D 形状的选中框），
且拖动连线时标签不跟随。**1-D 连接线自带文本**则只显示两端手柄、随线移动。

判据（读 vsdx 即可区分）：

```python
c = {x.get('N'): x.get('V') for x in shape.findall(NS+'Cell')}
one_d = 'BeginX' in c          # True = 1-D 连接线（自带文本）；False = 2-D 文本框
```

### 20.1 用法

```python
d.connect(b_srv, cp["srv"]["r44"], b_sw, cp["sw"]["l44"],
          label="双 10GE", label_color="gray", label_pos=(60, 49.5))
```

- `label_pos=(x, y)` 是**文字中心**的页坐标（mm，左上原点、Y 向下）；
  省略则取折线包围盒中心。
- `label_align` 默认 1（居中）、`label_size` 默认 7pt、`label_color` 默认随线色。

> ★★ **`label_pos` 已废弃，别再用**（2026-10-13 定案，见 §23）：绝对页坐标**不随线移动**，
> 一改布局标签就集体漂离线。改用 **`label_at` / `label_gap` / `label_side`** 沿线相对定位：
>
> ```python
> d.connect(b_srv, cp["srv"]["r44"], b_sw, cp["sw"]["l44"],
>           label="双 10GE", label_color="gray",
>           label_at=0.5, label_gap=2.2, label_side="up")
> ```

### 20.2 坑一：写 `Control.TextPosition` 会让文字**跳位并随线旋转**

Visio 原生「动态连接线」确实带 `Control.TextPosition` 命名行，但用 COM 按**形状局部坐标**
写入时，Visio 把它当**页面坐标**解释：

```python
# ✗ 不要这样写
r = c.AddNamedRow(9, "TextPosition", 0)
c.CellsSRC(9, r, 0).FormulaU = "%g mm" % tx     # tx 是局部坐标 → 被当页面坐标
c.CellsSRC(9, r, 1).FormulaU = "%g mm" % ty
```

症状：标签跳到别处、且方向随连接线走向旋转。**只设 `TxtPin/TxtLocPin/TxtAngle` 即可**。

### 20.3 坑二：短标签被挤成「一字一行」

连接线文本块**默认带左右页边距**（约 1.27mm/侧）→ 实际可用宽度比 `TxtWidth` 小 ~2.5mm。
实测：「授时」`TxtWidth=7.34mm`（文本实需 4.94mm）仍**逐字换行**；「双 10GE」`TxtWidth=11.73mm`
（实需 8.97mm）却单行 —— 差值 2.5mm 就是页边距。

```python
self._f(c, "LeftMargin", "0 mm")      # ★ 必须清零
self._f(c, "RightMargin", "0 mm")
self._f(c, "TxtWidth", "%g mm" % tw)  # 另留 4.5mm 余量（原 2.4mm 不够）
```

### 20.4 坑三：**不要**为了「多救几个标签」放宽 `_merge_line_labels` 的启发式

曾把 `max_w` 由 44 放宽到 60、把距离判据改成「框内离折线最近点」、并让「约束不一致」的
分组回退成连线标签候选。结果 `best.Text = t` 是**覆盖**写入 —— **多个文本框并到同一条
连接线时，后面的覆盖前面的**，图 8-1 的「输出文件复制回本机（回导）」整条文字**丢失**。

★ 正确做法：**要连线自带文本就显式传 `label=`**（§20.1），不要靠启发式猜。
启发式保持原状（`max_w=44.0`、中心距判据、不回退）。

### 20.5 落地检查清单

1. 跑图脚本 → 输出 `[文字溢出] 无`。
2. 读 vsdx：每个标签所在形状 `1D=True`。
3. 目视 PNG：**单行、水平**、落位与 `label_pos` 一致。
   ★ 量图别用「PNG 宽 / 页面宽」当比例 —— `Page.Export` 按**内容包围盒**裁剪，
   且包围盒用的是 `Pin ± Width/2`（折线形状的 LocPin 常不在中心）→ 比例会算错。
   稳妥做法：从已知图元的像素跨度反推 px/mm。
4. 重跑全量管线 → `verify_final` 全绿（图幅偏差 ≤0.1%）。
5. **回归另 4 张用同一库的图**（PNG 逐字节比对；`page1.xml` 逐字节比对；
   EMF 只看 `rclFrame` 的**宽高比**，绝对尺寸受会话/设备影响、不可比）。

---

## 21. ★★★ 折线选中时出现**矩形外框** —— 根因是 `ObjType`（2026-10-13 定位并修复）

**症状**：选中我们生成的折线连接线时，**沿包围盒画出一圈矩形外框**；
而手工在 Visio 里画的折线（原生 Dynamic connector）选中时**只有端点圆点 + 折点小方块**，
没有任何矩形。用户明确报「折线选中时依然出现了矩形框，和手绘折线效果不一致」。

**★ 根因 = `Misc.ObjType`**（单变量 A/B 逐项二分，9 组对照）：
`ObjType` 默认 `0`（由 Visio 自行判断）→ 形状被按普通 2-D 形状渲染，选中即画包围盒。
原生 Dynamic connector 是 **`ObjType = 2`**（`visLOFlagsRoutable`，
「可布线，必须是 1-D 形状」，见 MS 文档 *ObjType Cell, Miscellaneous Section*）
→ 选中只显示端点圆点 + 折点小方块。

| 变体（几何完全相同） | 选中时 |
|---|---|
| V0 基线（`ObjType=0`） | **矩形外框** ❌ |
| V1 **仅**加 `ObjType=2` | **无矩形** ✅ |
| V2/V3/V4/V5/V7（`+LockCalcWH` / `+LockHeight` / `+ConFixedCode` / `+LayerMember` / `+TxtPin`，仍 `ObjType=0`） | 都有矩形 ❌ |
| V8 全加（含 `ObjType=2`） | 无矩形 ✅ |
| REF Visio 原生 Dynamic connector（`ObjType=2`） | 无矩形 ✅ |

**⚠ 与 ShapeTransform 无关（早期结论有误，已推翻）**：
`LocPin=包围盒中心` 与 `LocPin=包围盒角` 两种写法**都出矩形**
（放大复核 `th_T1_old_zoom.png` / `th_T2_new_zoom.png`，两者都有框，只是手柄数量不同）。
`PinX=BeginX + 常量`（中心）写法仍保留 —— 它与 Visio 原生约定一致，无害。

### 正解（已落地 `visio_lib._author_path`）：**顺序敏感，务必照抄**

```python
# ① 先按「包围盒中心」把变换 + 几何写对
self._write_geom(conn, [(lx - dx0, ly - dy0) for lx, ly in loc])
# ② 先锁死布线，再认作「可布线 1-D 形状」
self._f(conn, "ConFixedCode", "3")
self._f(conn, "ObjType", "2")
# ③ 按 Visio 接管后的「真实」局部原点重写几何
ox = conn.CellsU("PinX").Result("mm") - conn.CellsU("LocPinX").Result("mm")
oy = conn.CellsU("PinY").Result("mm") - conn.CellsU("LocPinY").Result("mm")
self._write_geom(conn, [(bx + lx - ox, by + ly - oy) for lx, ly in loc])
```

**三个必须知道的副作用**：

1. **设 `ObjType=2` 会触发 Visio 布线引擎立即接管**：把 ShapeTransform 重写成原生约定
   `PinX=GUARD((BeginX+EndX)/2)`、`LocPinX=GUARD(Width*0.5)`、
   `Width=GUARD(EndX-BeginX)`、`Height=GUARD(EndY-BeginY)`，并**丢弃自写几何**
   （实测 5 点被换成 3 点自动路由）→ 必须**在其后**重写几何。
2. **`ConFixedCode` 必须写在 `ObjType` 之前**。顺序反了，Visio 会在随后的重算里
   **再次重布**：实测同一条横线，`ObjType→写几何→ConFixedCode` 得 6 点自动路由；
   `ConFixedCode→ObjType→写几何` 得精确 4 点。
3. **Visio 会把不足 0.25 in 的边夹取到 ±0.25 in**（横线 Height、竖线 Width、零长线两向）
   → 局部原点 `(PinX−LocPinX, PinY−LocPinY)` **≠ Begin 点**（实测横线偏 3.175 mm）。
   故**必须回读原点**，不能按 Begin 推算。

### 短形状 / 自由端会**在存盘后被改坏** → 自愈回退

`ObjType=2` 对**很短**的连接线（图 N10 时间刻度线 2.7 mm、箭头 5 mm）与
**有一端是自由点**的线（图 N8 引出线）会被 Visio 的 1-D 最小尺寸夹取改坏
（刻度线实测被缩到 0.475 mm；图 N8 被加出斜段）。
★ 这种改坏**发生在 `SaveAs` 之后**，写入当时读回是正常的 → **当场校验抓不到**。

`visio_lib.connect()` 的两道防线：

```python
free_end = isinstance(ia, tuple) or isinstance(ib, tuple)
if free_end:
    self._author_path(c, via, objtype=False)      # ① 自由端直接不加 ObjType
elif not self._author_path(c, via, objtype=True):
    c.Delete(); c = build()                        # ② 当场校验不合格 → 删除重建
    self._author_path(c, via, objtype=False)
```

代价：这些线选中时仍显示包围盒（图例线 / 刻度 / 箭头 / 总线引出）。
**图形正确优先** —— 渲染必须逐字节不变。

### 验收要点（本项目实测，2026-10-13）

- 12 张图重跑：**10 张 PNG 逐字节相同**，另 2 张**差异像素 0**；
  12 张 EMF `rclFrame` 纵横比偏差 **0.000%**。
- `ObjType=2` 共 **79 条**连接线（= 全部两端胶水的连接线）；
  `ObjType=0` 52 条（= 全部自由端线）。自愈回退触发 **0 次**。
- 交付件内嵌 `word/embeddings/*.vsdx` 中 `ObjType=2` 也是 **79 处**
  → 确认 OLE 内容已随源刷新，不是嵌的旧图。

### 怎么截「选中态」做验证

用 `PrintWindow(hwnd, hdc, 2)`（`PW_RENDERFULLCONTENT`）—— 它能渲染**被遮挡**的窗口，
比 `ImageGrab` 可靠；配合 `win.CenterViewOnShape(shape, 0)` 把目标居中，
再按**客户区中心**（`GetClientRect` + `ClientToScreen`；`PrintWindow` 输出是窗口坐标）裁剪。
★ `IVWindow` 没有 `CenterViewOn`，只有 `CenterViewOnShape(Shape, Flags)`。

---

## 22. ★ 图例 `legend()` 忽略 `w` → 整行越出画布，导出后右侧凭空多出一大片空白（2026-10-13 实坑）

**症状**：图上「主绘图区」显得偏窄，右侧一大片空白；图例各条目间距悬殊、最后一条挤到很远。
用户原话：「图主要区域 width 比图例小，导致右侧很多空白区域」。

**根因**：`legend()` 的 `w` 形参**从未被使用** —— 每个条目的文字框**硬编码 46mm**：

```python
self.text(cx + lw_ + 1.5, y - 2.6, 46, 5.2, label, ...)   # ← 46 写死
cx += lw_ + 1.5 + 46 + gap                                # 每条 +79.5mm
```

3 个条目从 `x=6` 起 → 终点 ≈ `6 + 3×79.5 = 238.5mm`，**远超 170mm 画布**。
而 `Page.Export` 是**按内容边界裁剪**（不是按页面）→ 把越界的图例也裁进图里，
于是主绘图区（6..166）只占画面左侧 ~2/3，右侧 1/3 是「图例文字 + 空白」。

**修法**：让 `legend()` 真正按 `w` **均匀分布**（整行自 `x` 起、到 `x+w` 止）：

```python
n = len(items)
cell = (w - gap * (n - 1)) / n          # 每格宽
tw   = max(6.0, cell - lw_ - 1.5)       # 文字框宽（不再写死 46）
# 每条：线样 cx..cx+lw_，文字 cx+lw_+1.5，然后 cx += cell + gap
```

传 `w <= 0` 时退化为旧的固定顺排（线样 26 + 文字 46 + gap）。

**向后兼容（重要）**：`图8_1` 原先传 `w=88` —— 小于它的自然宽度 153，
但旧版忽略 `w` 所以无碍。改成新语义后必须把 `w` 改成它的**真实占用宽 153**，
否则图例会缩到 6..94、文字换行。实测改后 **PNG 与 `page1.xml` 的 sha1 均逐字节不变**。
`图9_1` 改传 `w=160`（与内容区 6..166 同宽）。

**排查手法**：解析 vsdx，算每个形状的 `PinX ± Width/2`（单位是**英寸**，×25.4 得 mm），
看最右形状是否超过「画布宽 − 边距」。超了多半就是图例 / 长文本框越界，
而不是主绘图区真的窄 —— 别被「按内容裁剪」的导出骗了。

## 23. ★★ 连接线标签「离老远」—— 通用病：绝对坐标 + 包围盒中心兜底（2026-10-13 实坑）

**症状**：用户原话「9.1 连接线上的文本离连接线距离太远，**这可以是一个通用问题**」。
不是 图 9-1 的偶发，而是两条通用路径都有病。

**根因（三条）**：

1. **`label_pos=(x, y)` 是绝对页坐标，不随线移动。** 改一次布局（列宽 / 间距）
   全部标签集体漂离线 —— 图 9-1 实测偏出 **11~14mm**，还压到邻框、跨过泳道虚线边框。
2. **`label_pos=None` 的兜底是取折线「包围盒中心」。** 肘形线的包围盒中心常常
   **根本不在线上**（4 段肘线的中心往往落进空白三角区）。
3. **并入路径**（`_merge_line_labels`）把标签钉在**作者摆放的位置**，原样保留偏移：
   `VISIO_MERGE_DEBUG=1` 实测 fig_n8「生成/供」**5.0/5.5mm**、fig_n9 三条 **5.0mm**、
   fig_n10「周期开始/时间 →」**3.5mm** —— 看着像一段独立说明，而不像连接线标签。

**修法 A —— `connect()` 显式标签：改成「沿线相对定位」**（★ 推荐，永远别手写绝对坐标）：

```python
d.connect(a, cpa, b, cpb, label="组播")                 # 默认：最长段中点 + 横→上/竖→左 偏 2.2mm
d.connect(a, cpa, b, cpb, label="组播", label_at=(2, .5))  # 指定第 3 段中点
d.connect(a, cpa, b, cpb, label="组播", label_at=0.35)     # 全线弧长 35% 处
d.connect(a, cpa, b, cpb, label="组播", label_side="down", label_gap=2.6)
```

| 参数 | 含义 |
|---|---|
| `label_at=None` | 自动取**最长一段**的中点（最稳：线段最长、两侧最空） |
| `label_at=float` | 按整条折线**弧长比例**取点（0=起点，1=终点） |
| `label_at=(i, f)` | 第 `i` 段（0 起）内按比例 `f` 取点 —— **绕开障碍物**用它 |
| `label_gap` | 文字**中心**到线段的垂距 mm（默认 2.2） |
| `label_side` | `None/'auto'` 按走向自选（横线段→**上**、竖线段→**左**）；`'up'/'down'/'left'/'right'/'on'` 强制 |
| `label_pos` | 绝对页坐标，**不随线移动**；仅确需精确定位时用，用后必须肉眼看成品 |

锚点取在**线段上**、再沿法向偏 `gap`，因此「文字↔线」距离恒定、且随连接线一起动；
布局改了标签自动跟着走。

**修法 B —— 并入路径：贴线归一**。并入时**保持「沿线位置」和「左右侧」不变**，
只把「**字形**到线的距离」收敛到 `LBL_GAP = 0.8mm`：

```python
nd, si, npt, dv = _poly_near(bpts, cx, cy)     # 最近点 + 所在段方向
u = (cx-npt)/bd                                # 原来偏向哪一侧，就还偏哪侧
reach = a*b / sqrt((b*ux)^2 + (a*uy)^2)        # 轴对齐块沿 u 的中心→近边距
cx, cy = npt + u * (LBL_GAP + reach)           # 字形近边离线 0.8mm
```

字形盒 `(gw, gh)` 由源文本框的字号 + 行数估（`_lh` 行高 + PIL 字宽度量），
不是用文本框的 `w/h`（含左右页边距，会把标签推得过远）。

**两条配套经验（改完必须肉眼看成品）**：

1. **竖线段的 `label_gap` 必须 > 半字宽**：2 字标签半字宽 ≈ 2.47mm，默认 2.2 会
   **蹭到字右缘 0.3mm**（图 9-1「授时」实测）；4 字标签（「授时信号」）要 3.0mm 以上。
2. **自动侧会撞邻块 / 边框**：图 9-1「组播」放竖段左侧 → 跨过泳道虚线边框 x=122，
   改用 `label_at=(2, .5)` 落到**末段横线上方**才避开。走廊窄时就换段、换侧，
   别硬调 `label_gap`。

**排查手法**：
- 并入路径：`VISIO_MERGE_DEBUG=1` 会打印每个线标签的「原始距离 → 贴线后距离 + 字形尺寸」。
- 显式路径：解析 vsdx 读该形状的 `TxtPinX/TxtPinY`，配合 `PinX−LocPinX` 反算文字中心，
  与 `_line_page_pts()` 的折线点列比距离（超 3mm 就是漂了）。

**验收（本项目 12 图回归）**：仅 fig_n8 / fig_n9 / fig_n10 与 图9_1 变化，
差异像素 0.126% / 0.073%（集中在标签本身），其余 8 图零变化；
合并标签 5.0/5.5/3.5mm → 全部收敛到「字形离线 0.8mm」。


## 24. ★★★ 连接线从框的**错误一侧**出线 —— 会横穿整个框，肉眼极难发现（2026-10-17 实坑）

**症状**（用户原话）：「『比对（GUI）→ 报表』的连接线**起始位置不对，跨越了整个比对（GUI）**，
此问题应按**共性问题**排查」。

**根因**：`connect()` 的 `ia/ib` 只是「胶到哪个连接点」，**不校验方向是否合理**：

```python
cp["cmp"] = d.side_cps(box["cmp"], "bottom left")
d.connect(box["cmp"], cp["cmp"]["left"], box["rep"], cp["rep"]["left"])   # ✗
#                     ^^^^^^^^^^^^^^^ 目标在右边，却从左边出线
```

两点同高时，连线**从左边中点一路向右**，横穿整个源框。源框有底色 → 线被压在框下面，
只在右侧缺口露一截箭头，成品看着像「从框左边划过来」，**源脚本读起来完全正常**。
另一形态：终点胶在**左边中点**、路径却绕进框内再横穿到边缘（本项目 图N11 记录框）。

**为什么必须当共性问题查**：同一份脚本里 `side_cps(...)` 给的侧别，很容易与实际布局
（谁在左、谁在右）不一致；改一次布局就可能出现，且**不会报任何错**。

**修法（两层，都已内置）**：

1. **写的时候就拦**：`connect()` 现在会做**端点出线方向自检** —— 该端连接点的
   朝外法线 · 首段方向 < 0 时立即打印：

   ```
   [连线方向] Begin 端胶在形状#18 的 left 侧，首段却朝反方向 → 连线会横穿该形状。
              改用朝向目标的那一侧（例：connect(a, cp["right"], b, cp["left"])）。
   ```
   只查**四向中点**（自定义连接点无「朝向」语义，跳过）；自由点 / 中心胶水跳过；
   先向左绕再回头的合法走法**不会误报**（法线与首段同向）。

2. **`d.check()` 现在连整页一起查**：除文字溢出外，还输出

   ```
   [连线穿越] 线#7 穿入 形状#1（30×18 @10.0,10.0）内部
   [连线自检] 出线方向 1 处 ｜ 穿越实体框 1 处
   ```
   `audit_links()` 用 Liang-Barsky 判「折线是否穿入某个**有底色**框内部」，
   并排除两类「穿过去也正常」的框：
   - **容器**：面积 ≥18% 页面积（泳道 / 系统边界框），或内含 ≥2 个**有底色**子框；
   - **无底色文本框**（`FillPattern=0`）：通栏标题框，线从空白处经过不压字。

   ★ 容器判定里「子框只数有底色的」是关键：`step_box` 内部还有 1~2 个无底色文字框
   （`merge_texts` 之前仍在页面上），若一并计数，**每个 step_box 都会被误判成容器**
   而整张图漏查。

**自查口径（解析 vsdx，不启 Visio）**：

- `Cell` 的 `V` 是**内部英寸**（`U='MM'` 只是显示单位）→ 当毫米用差 25.4 倍；
- 连接点段名在文件里是 **`Connection`（单数）**，`Connects` 却引用 `Connections.Xn`；
- **`Connections.Y` 原点在左下、Y 向上** → 自上起算 `fy = 1 − Y/Height`（写反则上下全错）；
- `pages.xml` 的属性是**单引号**，正则要 `['"]`。

**回归用例（改完自检逻辑必跑）**：

| 用例 | 期望 |
|---|---|
| 左→左（目标在右） | 报「出线方向」+「穿越实体框」 |
| 右→左（正常） | 无 |
| 上→下，中间隔一个框 | 只报「穿越实体框」（**不应**报出线方向） |
| 上→下，`via` 绕开中间框 | 无 |
| 先向左绕行再回头接右边 | 无（合法走法不误报） |

**本项目全量结果**：15 张图 2 处（图 8-1「比对→报表」、图N11「记录框」），其余 13 张 0 处；
修完 `d.check()` 恒为「均无」。
