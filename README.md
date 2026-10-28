# visio-diagram

[中文文档 (Chinese)](README.zh-CN.md)

Programmatic vector diagram generation and Word embedding for Chinese engineering documents, driven by the local Visio installation via COM.

Each figure is exported as `.vsdx` (natively editable), `.emf` (vector for Word / PowerPoint) and `.png` (preview).

## Capabilities

- Precise layout control: rectangles, text blocks, titled boxes, swimlanes, legends, horizontal / vertical arrows
- Single-shape polyline connectors glued to connection points — endpoints follow when boxes move
- Connector-carried labels positioned along the line (`label_at` / `label_gap` / `label_side`)
- Text overflow self-check before export (font-metrics based line-wrap estimation)
- Arrow alignment audit: flags endpoints that land off the edge midpoint; qualified offsets can be whitelisted
- Automatic text merging into host shapes before save (≈47% fewer shapes, pixel-identical layout)
- Word embedding: placeholder replacement (EMF), raw OOXML assembly, or in-place OLE conversion (double-click to edit in Visio)
- Document QA: static OOXML checks, Word layout measurement, regression suite, docs-vs-code consistency check

## Requirements

- Windows with desktop Visio installed
- Python 3.9 – 3.13
- `pywin32`, `Pillow`
- A CJK font (default: Microsoft YaHei)

## Installation

```bash
git clone https://github.com/kkzi/visio-diagram.git
cd visio-diagram
pip install pywin32 pillow
```

Verify the environment:

```bash
python -c "import sys; sys.path.insert(0,'scripts'); import visio_lib; visio_lib.check_env()"
```

## Install as an LLM agent skill

The repo is a standard skill package (`SKILL.md` + `scripts/` + `references/`). Clone it into any agent's skill directory and the agent can drive the full workflow autonomously:

```bash
# WorkBuddy (user-level)
git clone https://github.com/kkzi/visio-diagram.git ~/.workbuddy-ai/skills/visio-diagram

# Claude Code
git clone https://github.com/kkzi/visio-diagram.git ~/.claude/skills/visio-diagram
```

Example prompt that triggers the skill:

```
用 Visio 画一张三泳道的系统架构图，按文档栏宽 1:1 出图，自检后嵌到方案 docx 里
```

The agent then runs the workflow itself: environment check → column-width measurement → drawing script → overflow self-check → alignment audit → export and Word embedding. The host machine needs the same requirements as above; the agent only needs shell access.

## Usage

```python
import sys
sys.path.insert(0, "scripts")
from visio_lib import VDraw

with VDraw(150, 108, theme="tech-cn") as d:      # canvas in mm, top-left origin, Y down
    d.lane(0, 0, 150, 74, "处理链路")
    d.step_box(6, 12, 42, 26, "① 环节一", "说明文字")
    d.h_arrow(48, 25, 60)
    d.check()                                     # text overflow self-check
    d.export_all("out/图N1")                      # writes .vsdx / .emf / .png
```

Embed a figure into Word by placeholder:

```bash
python scripts/word_insert.py 方案.docx --placeholder "{{FIG:N2}}" --image 图N2.emf
```

Audit connector alignment (supports a whitelist for intentional offsets):

```bash
python scripts/audit_align.py 图表素材/*.vsdx
python scripts/audit_align.py --add 图表素材/图N8….vsdx    # whitelist reported offsets
```

Run the regression suite after any engine change:

```bash
python scripts/selftest.py          # 19 assertions
python scripts/doc_consistency.py  # docs-vs-code consistency
```

Full reference: [SKILL.md](SKILL.md) and [references/](references/).

## Known Limitations

1. Free-end connectors (legend lines, ticks, bus stubs) show a selection bounding box; graphical correctness is prioritized.
2. Automatic text merging is heuristic; use explicit `label=` for precise control.
3. Requires local Visio, pywin32 and CJK fonts — not portable to machines without them.
4. COM sessions must be closed; use `with VDraw(...)` so `VISIO.EXE` never leaks.

## License

MIT — see [LICENSE](LICENSE).
