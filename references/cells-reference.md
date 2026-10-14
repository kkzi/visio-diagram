# Visio CellsU 速查与坐标换算

## 一、坐标换算（最容易出错的地方）

Visio 内部：**原点在页面左下角，Y 轴向上**，默认单位英寸。
本 skill 的 `visio_lib` 对外统一为 **左上角原点、Y 向下、mm**，换算只有两行：

```python
PinX = x + w / 2.0                    # 水平方向两者一致
PinY = pageH - ytop - h / 2.0         # 垂直方向要翻转
Width  = w
Height = h
```

直线（`DrawLine`）同理，只换 Y：

```python
BeginY = pageH - y1
EndY   = pageH - y2
```

**写 Cell 时一定要带单位**，否则 Visio 按英寸解释：

```python
s.CellsU("PinX").FormulaU = "25 mm"      # ✓
s.CellsU("PinX").FormulaU = "25"         # ✗ 会被当成 25 英寸
s.CellsU("Char.Size").FormulaU = "8.5 pt"   # 字号用 pt
```

`CellsU` 取值一律是 **字符串**，布尔用 `"0"` / `"1"`，不要传 Python 的 `True`。

---

## 二、PageSheet（页面）

| Cell | 说明 | 示例 |
|---|---|---|
| `PageWidth` / `PageHeight` | 页面尺寸 | `"152 mm"` |
| `LeftMargin` / `RightMargin` / `TopMargin` / `BottomMargin` | 页边距，**设为 0** 便于按内容排布 | `"0 mm"` |

```python
pg = doc.Pages.Item(1)
pg.PageSheet.CellsU("PageWidth").FormulaU = "152 mm"
```

---

## 三、Shape（图元）

### 几何
| Cell | 说明 |
|---|---|
| `PinX` / `PinY` | 形状**中心**坐标 |
| `Width` / `Height` | 尺寸 |
| `BeginX/BeginY/EndX/EndY` | 仅线条类形状（`DrawLine`） |
| `Rounding` | 圆角半径，`"0 mm"` 直角，`"1.5 mm"` 圆角 |

### 填充
| Cell | 说明 |
|---|---|
| `FillPattern` | `0` 无填充，`1` 实心，2~24 为各种纹理 |
| `FillForegnd` | 填充色，`"RGB(222,235,247)"` |

### 线条
| Cell | 说明 |
|---|---|
| `LinePattern` | `0` 无线 `1` 实线 `2` 虚线 `4` 点划线（编号共几十种，见下） |
| `LineColor` | `"RGB(0,0,0)"` |
| `LineWeight` | `"0.75 pt"` |
| `EndArrow` / `BeginArrow` | 箭头样式编号 |
| `Shadows` | 一般设 `"0"` 关闭阴影 |

### 文字
| Cell | 说明 |
|---|---|
| `Char.Font` | 字体名，**必须带引号**：`'"微软雅黑"'` |
| `Char.Color` | `"RGB(38,38,38)"` |
| `Char.Size` | `"8.5 pt"` |
| `Para.HorzAlign` | `0` 左 `1` 中 `2` 右 `3` 两端对齐 |
| `VerticalAlign` | `0` 顶 `1` 中 `2` 底 |
| `LeftMargin` / `RightMargin` | 文字与框边的水平内边距 |
| `TopMargin` / `BottomMargin` | 垂直内边距 |

文字内容用属性赋值，不是 Cell：

```python
s.Text = "标题\n说明文字"        # 换行用 \n
```

> `Char.Font` 忘了加引号是最常见的静默失败：不报错，但字体不生效。

---

## 四、编号取值（以实际渲染为准）

Visio 的线型与箭头编号有几十种，各版本略有差异。**可靠做法**：在 Visio 界面里手动设一次，再读回 `CellsU(...).FormulaU` 看编号。

本 skill 验证过的常用值：

| 用途 | 值 |
|---|---|
| 无箭头 | `EndArrow = 0` |
| 实心三角箭头 | `EndArrow = 4` |
| 空心箭头 | `EndArrow = 1` |
| 实线 | `LinePattern = 1` |
| 虚线 | `LinePattern = 2` |
| 点划线 | `LinePattern = 4` |

---

## 五、创建图元的固定套路

Visio COM 没有「按矩形创建一个矩形」的直接 API，套路是**先建默认图元再摆位**：

```python
s = pg.DrawRectangle(0, 0, 1, 1)     # 先建一个 1x1 的
s.CellsU("PinX").FormulaU = "..."    # 再改尺寸和位置
s.CellsU("PinY").FormulaU = "..."
s.CellsU("Width").FormulaU  = "60 mm"
s.CellsU("Height").FormulaU = "20 mm"
```

线条用 `pg.DrawLine(0, 0, 1, 1)`，然后改 `BeginX/BeginY/EndX/EndY`。

---

## 六、导出与保存

```python
doc.SaveAs(r"D:\out\图A.vsdx")     # 原生格式，可继续编辑
pg.Export(r"D:\out\图A.emf")       # 矢量，供 Word/PPT 嵌入
pg.Export(r"D:\out\图A.png")       # 位图预览
```

- `Page.Export` **按内容边界裁剪**（不含页边距），96 DPI —— 所以不用手工裁白边。
- 扩展名决定格式，由 Visio 的导出滤镜处理。
- `Export` 只导出当前页；多页文档要对每页分别调用。

---

## 七、COM 生命周期

```python
app = win32.Dispatch("Visio.Application")
app.AlertResponse = 7      # 抑制弹窗（否则可能卡住等待人工点击）
app.Visible = 0
doc = app.Documents.Add("")
...
doc.Close()
app.Quit()
```

**务必 `Quit()`**，否则 VISIO.EXE 会残留在后台进程里，多次运行后累积。
