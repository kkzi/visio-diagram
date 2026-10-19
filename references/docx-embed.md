# 把 Visio 图嵌进 Word

三条路线：**A**（插图，静态矢量）→ **C**（OLE 对象，双击可编辑）→ **B**（手工拼 OOXML，兜底）。
选哪条取决于「甲方要不要在 Word 里双击改图」：

| 需求 | 路线 | 产物 |
|---|---|---|
| 只要图能看、缩放不糊 | A / B | `<w:drawing>` + EMF |
| 甲方要**双击进 Visio 编辑** | **C** | `<w:object>` + `<o:OLEObject ProgID="Visio.Drawing.15">` |
| 无 Word 环境，纯脚本合成 | B | 同 A |

**路线 C 建议叠在 A 之上做**：先用 A/B 把 EMF 图插进去（排版、宽度都调好），
最后再用 C 把每张图**原地**换成 OLE 对象——这样排版逻辑完全复用，不必重算尺寸。

---

## 路线 C：转成可双击编辑的 Visio 对象（OLE）

甲方常见要求：「图要能在 Word 里双击打开 Visio 改」。这时不能插 EMF，要嵌 **OLE 对象**。

**做法**：Word COM 的 `InlineShapes.AddOLEObject()`，**原地替换**已有图片，保留原显示宽度。

```python
import win32com.client
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False; word.DisplayAlerts = 0
doc = word.Documents.Open(DOC, ReadOnly=False)
shapes = list(doc.InlineShapes)                 # 必须先快照，替换过程中集合会变
assert len(shapes) == len(FIGS), "图片数与预期不符"
for sh, fig in zip(shapes, FIGS):
    rng, w, h = sh.Range, sh.Width, sh.Height   # 先记住 Range + 宽 + 高
    sh.Delete()                                 # 再删，否则 Range 失效
    new = rng.InlineShapes.AddOLEObject(FileName=abs_vsdx)
    new.LockAspectRatio = 0                     # msoFalse：必须先关
    new.Width = w                               # 宽度还原
    new.Height = h                              # ★ 高度也必须还原（见坑 3）
for toc in doc.TablesOfContents:
    toc.Update()                                # 图数变了，页码要重排
doc.Save(); doc.Close(False); word.Quit()
```

**要点 / 坑**：

1. **顺序敏感**：`InlineShapes` 按文档出现顺序排列，`FIGS` 必须与之一一对应。
   先 `list()` 快照再改，边遍历边删会漏项。
2. **先取 `Range` 再 `Delete`**：删掉形状后 `sh.Range` 就失效了，插入点要用删之前的 `rng`。
3. **★ 宽度和高度都要手动还原**：`AddOLEObject` 既不继承原宽度、也不继承原高度 ——
   它按 **OLE 对象（Visio）自身固有纵横比**定尺寸。只写 `Width` 时 Word 会据此重算高度，
   结果与 EMF 图幅不符，**图被压扁（实测最高 +20%）**，而页数/表格/禁词等常规校验全绿、查不出来。
   所以必须 `LockAspectRatio = 0` 后把 `Width` 与 `Height` **都**写回（原高度取自替换前的图，即 EMF 比例）。
   详见 `pitfalls.md` §13。
4. **嵌的是 `.vsdx`**，不是 `.emf`——所以源文件要留 `.vsdx`（路线 A/B 只需 `.emf`）。
5. **TOC 必须 `Update()`**：图片换成 OLE 后页面重新分页，目录页码会全错。
6. **完成后要去掉 `settings.xml` 的 `<w:updateFields/>`**，否则甲方一打开就弹「是否更新域」。
7. **Word 会把 EMF 预览重写一遍**（`word/media/*.emf` 字节数会变），这是正常的；
   校验时不要拿嵌入 EMF 的 sha1 去和源 EMF 比，要比 **`.vsdx` 的对象结构**。

**校验**（换完后必查）：

```python
d = zipfile.ZipFile(docx).read('word/document.xml').decode('utf-8')
d.count('<w:object')          # == 图数
d.count('Visio.Drawing.15')   # == 图数
d.count('<w:drawing')         # == 0（不应再有裸图）
len([n for n in z.namelist() if 'embeddings/' in n])   # == 图数
len([n for n in z.namelist() if 'media/' in n])        # == 图数（EMF 预览）

# ★ 图幅校验：v:shape 的 style 宽高比须与 EMF 宽高比一致（偏差 ≤0.1%）
import re
from docx_embed import emf_info
styles = re.findall(r'style="width:([0-9.]+)pt;height:([0-9.]+)pt"', d)
for (wp, hp), emf in zip(styles, EMFS):        # EMFS 顺序同 FIGS
    assert abs(float(wp) / float(hp) / emf_info(emf)["ratio"] - 1) < 0.001
```

再开一次 Word 读 `doc.InlineShapes.Count` 与 `Type==1`（wdInlineShapeEmbeddedOLEObject）的个数，确认全部是 OLE。

**注意**：Word 保存后会**规范化 XML**（如 `<w:b />` → `<w:b/>`、属性顺序变化）。
用字符串计数做校验时，两种写法都要匹配，别只按一种写。

---

## 路线 A：Word COM 插入（推荐）

Word 原生支持 EMF，矢量保真，**不用碰 OOXML**。

**做法**：在文档里留占位符，画完图后批量替换。

```markdown
... 系统工作模式见下图。

{{FIG:N2}}

图 9-2 系统工作模式与数据流向图
```

```bash
python scripts/word_insert.py 方案.docx --placeholder "{{FIG:N2}}" --image 图N2.emf
```

脚本会：找到所有 `{{FIG:N2}}` → 清空占位符 → 在插入点插入图片 → 按正文栏宽设宽度 → 按 EMF 宽高比设高度 → 段落居中 → 更新域 → 另存。

默认宽度 = **正文栏宽 − 5 mm**（`content_column_mm()` 自动算出，已计入装订线）。

也可当库用：

```python
from word_insert import insert_at_placeholder
insert_at_placeholder("in.docx", "{{FIG:N3}}", "图N3.emf", width_mm=170, out="out.docx")
```

**优点**：不依赖文档生成方式；EMF 尺寸由 Word 自己解析；矢量不丢。
**前提**：本机装了 Word。

---

## 路线 B：手工拼 OOXML

适用于「用脚本从零合成 docx」的场景（无 Word 或需要完全程序化）。

### B1. EMF 图幅要从文件头读

PIL **读不了 EMF**（会抛异常），`python-docx` 的 `add_picture` 内部用 PIL，所以**无法直接插入 EMF**。必须自己从 EMF 头取尺寸：

```python
from docx_embed import emf_info
info = emf_info("图N2.emf")
# {'bounds_px': (523,539), 'frame_mm': (113.4,116.8),
#  'ratio': 1.0302, 'implied_dpi': 117.2, ...}
```

EMF 头结构（ENHMETAHEADER）：

| 偏移 | 字段 | 说明 |
|---|---|---|
| 0 | `iType` | 恒为 1 |
| 4 | `nSize` | 头长度 |
| 8 | `rclBounds` | 4×int32，**设备单位（像素）** |
| 24 | `rclFrame` | 4×int32，**单位 0.01 mm** |
| 40 | `dSignature` | 应为 `" EMF"`（`0x464D4520`） |

写入 `wp:extent` 时换算成 EMU：`cx = mm * 36000`。

> ⚠ **EMF 的绝对物理尺寸不可靠**。实测 Visio 导出的 EMF，`rclFrame` 隐含约 **117 DPI**，
> 而同一张图导出的 PNG 是 **96 DPI**（物理尺寸差约 22%）。
> 但两者的**宽高比完全一致**（差异 < 0.1%）。
>
> **正确做法**：只取宽高比，宽度按正文栏宽显式指定，高度 = 宽度 × `ratio`。
> 不要直接采信 `rclFrame` 的绝对值 —— 否则图会莫名缩小或放大。
> `word_insert.py` 与 `docx_audit.py` 都是这么做的。

### B2. 补 Content_Types 声明

```python
from docx_embed import ensure_content_type
ensure_content_type("方案.docx", "emf")     # 补 <Default Extension="emf" ContentType="image/x-emf"/>
```

### B3. 三个必踩的 OOXML 坑

#### 坑 1：`<Default>` 必须写在 `<Types>` 内

`[Content_Types].xml` 只有一个根节点 `<Types>`。如果把 `<Default>` 插到 XML 声明之后、`<Types>` 之前，文件就有两个根级元素 —— Word 报**「文件已损坏」**。

```python
m = re.search(r"<Types[^>]*>", ct)
ct = ct[:m.end()] + decl + ct[m.end():]      # ✓ 插在 <Types> 开标签之后
```

#### 坑 2：`mc:Ignorable` 只能引用已声明的前缀

根元素形如：

```xml
<w:document xmlns:w="..." xmlns:mc="..." mc:Ignorable="w14 wp14">
```

如果用了 `ElementTree` 解析再序列化，**未被引用的命名空间声明会被丢掉**，但 `mc:Ignorable` 还留着 `w14` —— 违反 MCE 规范，Word 同样报**「文件已损坏」**。

修复：序列化后重新裁剪 `mc:Ignorable`，只保留仍声明了的前缀。

```python
head = s[:s.find(">", s.find("<" + root_tag)) + 1]
declared = set(re.findall(r'xmlns:([\w\d]+)="', head))
keep = [p for p in ignorable.split() if p in declared]
```

`docx_audit.py` 已内置这项检查，对 `document.xml` 与 `styles.xml` 都会查。

#### 坑 3：`w:tblW` / `w:tcW` 的属性名是 `w:w`，不是 `w:val`

```xml
<w:tblW w:w="9921" w:type="dxa"/>     <!-- ✓ -->
<w:tblW w:val="9921" w:type="dxa"/>   <!-- ✗ -->
```

写成 `w:val` 时 Word 读到的 `PreferredWidth = 0`，**每一列塌缩到约 0.4 cm**（一行一个字），页数会从 32 页暴涨到 201 页。

容易混淆的是 `w:gridCol` 用的确实是 `w:w`（这个是对的），所以不是所有宽度属性都写错 —— 只有 `tblW` / `tcW` 这两个容易顺手写成 `w:val`。

`docx_audit.py` 已内置该检查。

---

## 交付前自查

```bash
python scripts/docx_audit.py 方案.docx
```

静态部分（不需要 Word）会检查：

- 所有 XML 部件良构
- `<Default>` 是否在 `<Types>` 内、`word/media` 的扩展名是否都已声明
- `mc:Ignorable` 前缀一致性
- `tblW` / `tcW` 属性名
- 加粗 run 计数
- 正文栏宽（含装订线）

Word 实测部分会给出：页数、段落、表格数、**最宽表格 vs 栏宽**、每张图的尺寸、表格内外粗体段落数。

判定口径：
- 最宽表格 ≤ 正文栏宽 → 通过（差 0.05 cm 容差）
- 表格内粗体 = 0 → 通过
- 表格外粗体 = 标题段落数 → 正常（标题样式自带加粗）

---

### 图的版式验收（图幅 / 居中 / 目录页码）

`docx_audit.py` 只看文本和表格，下面三项图特有的必须单独跑。三项都要在**导出 PDF 之后**做。

**① 图幅比例 —— 三方对照，别猜**

```python
# 显示框（docx）、bounds（EMF 头）、vsdx 并集（设计真值）三者比对
# 判据与决策树见 pitfalls.md §25
dr = dw / dh                     # <wp:extent> 或 <v:shape style>，注意 in/pt/mm 单位（§26）
br = (bd[2] - bd[0]) / (bd[3] - bd[1])   # struct.unpack('<4i', emf[8:24])
vr = vsdx_union(path)            # PinX/PinY/Width/Height 并集，单位英寸
```

判定：`|br/vr-1| ≤ 3%` 且 `|dr/vr-1| > 3%` → 显示框写错，改 `H = W / br`。
出报告时把 `dr / br` 的偏差列全表，**>1.5% 的都要看**（本项目首轮 25 个恒 −48%，
精修后 12 个残余 1.7%~3.4%，都收敛到 <1.5% 才算完）。

**② 水平居中 —— PDF 矢量路径法**

```python
rects = [d['rect'] for d in page.get_drawings()
         if d['rect'].width > 0.5 and d['rect'].width <= COLW * 1.06]   # 排除 EMF 超大裁剪框
# 按题注 y 区间取并集包围盒，与栏中心比对，偏差 ≤1pt
```

★ 不要用 `Range.Information(5)`（对行内图形一律返回段落左边距）；也不要按「墨迹包围盒」
切带（图内文字会被当成正文，导致假阳性）。细节见 `pitfalls.md` §28.1。

**③ 目录 / 图表目录页码**

改图尺寸必然改页数（本项目 204 → 201）。刷域之后**必须**重新导 PDF 逐条比对：
目录条目的页码 vs 正文题注实际所在页。

```python
doc.Fields.Update()
for i in range(1, doc.TablesOfContents.Count + 1):   doc.TablesOfContents(i).Update()
for i in range(1, doc.TablesOfFigures.Count + 1):    doc.TablesOfFigures(i).Update()
doc.Repaginate(); doc.Save()
```

页码要从**页脚块**读（`blocks` 中 `y0 > 页高-70` 的纯数字块），不能按 `get_text()` 行序取末尾
（见 `pitfalls.md` §28.2）。

**导出 PDF 的姿势**

```python
shutil.copy2(docx, tmp)                        # 走临时副本，原文档不动
doc = app.Documents.Open(os.path.abspath(tmp)) # ★ 必须 abspath，否则 Word 找 System32
doc.SaveAs2(os.path.abspath(pdf), 17)          # 17 = wdFormatPDF
```

`ExportAsFixedFormat` 在部分环境**静默失败**（不报错、不产文件），优先用 `SaveAs2`。
