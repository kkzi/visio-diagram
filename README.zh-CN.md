# visio-diagram

[English](README.md)

通过 COM 驱动本机 Visio，程序化生成矢量图并嵌入 Word 文档，面向中文工程技术文档配图。

每张图导出三件套：`.vsdx`（原生可编辑）· `.emf`（矢量，供 Word / PPT）· `.png`（预览）。

## 能力范围

- 精确版式控制：矩形、文本块、标题方块、泳道、图例、水平 / 垂直箭头
- 单条折线连接线，两端胶到连接点 —— 挪框时端点跟随
- 连接线自带标签，沿线相对定位（`label_at` / `label_gap` / `label_side`）
- 出图前文字溢出自检（基于字体度量的折行估算）
- 箭头对齐审计：报出偏离边中点的端点；已定性的刻意偏移可登记白名单
- 保存前自动把文本框并入宿主形状（形状数省约 47%，版式逐像素不变）
- 嵌入 Word：占位符替换（EMF）/ 手拼 OOXML / 原地换 OLE（双击进 Visio 改图）
- 文档自检：静态 OOXML 检查、Word 版式实测、回归测试、文档-代码一致性检查

## 环境要求

- Windows + 桌面版 Visio
- Python 3.9 – 3.13
- `pywin32`、`Pillow`
- 中文字体（默认微软雅黑）

## 安装

```bash
git clone https://github.com/kkzi/visio-diagram.git
cd visio-diagram
pip install pywin32 pillow
```

环境自检：

```bash
python -c "import sys; sys.path.insert(0,'scripts'); import visio_lib; visio_lib.check_env()"
```

## 作为 LLM agent 技能安装

本仓库即标准技能包（`SKILL.md` + `scripts/` + `references/`）。克隆到 agent 的技能目录后，agent 即可自主驱动完整工作流：

```bash
# WorkBuddy（用户级）
git clone https://github.com/kkzi/visio-diagram.git ~/.workbuddy-ai/skills/visio-diagram

# Claude Code
git clone https://github.com/kkzi/visio-diagram.git ~/.claude/skills/visio-diagram
```

触发技能的示例指令：

```
用 Visio 画一张三泳道的系统架构图，按文档栏宽 1:1 出图，自检后嵌到方案 docx 里
```

agent 会自行完成：环境检查 → 量栏宽 → 写绘图脚本 → 溢出自检 → 对齐审计 → 出图并嵌入 Word。宿主机环境要求同上；agent 只需能执行 shell 命令。

## 使用

```python
import sys
sys.path.insert(0, "scripts")
from visio_lib import VDraw

with VDraw(150, 108, theme="tech-cn") as d:      # 画布 mm，左上原点、Y 向下
    d.lane(0, 0, 150, 74, "处理链路")
    d.step_box(6, 12, 42, 26, "① 环节一", "说明文字")
    d.h_arrow(48, 25, 60)
    d.check()                                     # 文字溢出自检
    d.export_all("out/图N1")                      # 导出 .vsdx / .emf / .png
```

按占位符把图嵌入 Word：

```bash
python scripts/word_insert.py 方案.docx --placeholder "{{FIG:N2}}" --image 图N2.emf
```

连接线对齐审计（支持把已定性偏移登记白名单）：

```bash
python scripts/audit_align.py 图表素材/*.vsdx
python scripts/audit_align.py --add 图表素材/图N8….vsdx    # 登记本次报出的偏移
```

改完引擎先跑回归：

```bash
python scripts/selftest.py          # 14 项断言
python scripts/doc_consistency.py  # 文档-代码一致性
```

完整参考见 [SKILL.md](SKILL.md) 与 [references/](references/)。

## 已知限制

1. 自由端连接线（图例线 / 刻度 / 总线引出）选中时仍显示包围盒；本库取「图形正确优先」。
2. 自动文字并入是启发式；要精确控制请用显式 `label=`。
3. 依赖本机 Visio、pywin32 与中文字体 —— 无法在缺少这些依赖的机器上使用。
4. COM 会话必须收尾；用 `with VDraw(...)` 即可保证不残留 `VISIO.EXE`。

## 许可证

MIT —— 见 [LICENSE](LICENSE)。
