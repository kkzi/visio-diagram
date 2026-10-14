# -*- coding: utf-8 -*-
"""文档-代码一致性检查 —— 防止「结论漂移」。

背景：本 skill 的坑清单（`references/pitfalls.md`）、使用说明（`SKILL.md`）与引擎
（`scripts/visio_lib.py`）是三处**独立维护**的文件。历史上出现过「代码已改用
`ObjType`、文档仍写 `ShapeTransform`」这类**旧结论残留**（2026-10-13 一次就抓到
两处，且都在同一结论上）。本脚本把几条关键结论固化成**机器可查**的规则，改完库/文档
跑一遍即可。

规则三类：
  · FORBIDDEN —— 已废止的表述，出现即 FAIL（旧根因 / 旧做法）。
  · REQUIRED  —— 指定文件**必须**包含的关键词（结论没被误删）。
  · COUPLED   —— 文件里出现 A 就**必须**同时出现 B（结论成对出现，防只改一半）。

用法：
    python scripts/doc_consistency.py
退出码：0 = 全通过；1 = 有问题。**只读**，不改任何文件。
"""
import os
import sys

# ★ Windows 控制台默认 GBK，打印 ✅/❌ 会抛 UnicodeEncodeError → 退出码变 1（假失败）。
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)                     # skill 根目录

# ---------------------------------------------------------------- 规则表
# 已废止表述（全 skill 范围，出现即失败）。
# ★ 注意：pitfalls.md 会在「反例」里**有意引用**旧写法，所以这里只放
#   **纯结论性**的旧表述；脚本自身也要排除（见 SELF）。
FORBIDDEN = [
    # 2026-10-13 被推翻的旧结论：把「选中矩形外框」归因到 ShapeTransform
    "ShapeTransform 须用标准 1-D 约定",
    "LocPin 落在包围盒角上 → Visio 把它按可缩放 2-D 形状渲染",
]

# 定向禁用：只在**指定文件**里算违规。
# ★ `label_pos=(` 在 pitfalls.md §20.1 是**有意保留的旧用法反例**（§23 才给否定结论），
#   不能全局禁；但 SKILL.md 的「推荐写法」示例里出现它 = 自相矛盾（2026-10-13 评审 P3-6）。
#   正确写法是 label_at / label_gap / label_side 沿线相对定位。
FORBIDDEN_IN = {
    "SKILL.md": ["label_pos=("],
}

# 指定文件必须包含的关键词
REQUIRED = {
    "references/pitfalls.md": ["ObjType", "ConFixedCode", "回读", "legend", "label_at"],
    "SKILL.md": ["ObjType", "legend", "label_at"],
    "scripts/visio_lib.py": ["ObjType", "ConFixedCode", "label_at"],
}

# 成对约束：文件里出现 A 就必须出现 B
COUPLED = [
    (["矩形外框"], ["ObjType"], "提到「矩形外框」就必须同时说明真因是 ObjType"),
    (["老远"], ["label_at"], "提到「标签离线远」就必须同时给出 label_at 沿线定位的修法"),
]

TEXT_EXT = (".md", ".py")
SELF = "scripts/doc_consistency.py"              # 本脚本自带规则字串，跳过扫描


def iter_files():
    for base, dirs, files in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in ("__pycache__", ".git")]
        for f in files:
            if f.endswith(TEXT_EXT):
                p = os.path.join(base, f)
                rel = os.path.relpath(p, ROOT).replace("\\", "/")
                if rel != SELF:
                    yield rel, p


def main():
    problems = []
    files = list(iter_files())
    texts = {}
    for rel, p in files:
        try:
            with open(p, encoding="utf-8") as fh:
                texts[rel] = fh.read()
        except Exception as e:                       # pragma: no cover
            problems.append("[读取失败] %s: %s" % (rel, e))

    # 1) FORBIDDEN
    for rel, txt in sorted(texts.items()):
        for bad in FORBIDDEN:
            if bad in txt:
                problems.append("[废止表述] %s 仍含：%s" % (rel, bad))

    # 1b) FORBIDDEN_IN（定向：只在该文件里算违规）
    for rel, bads in FORBIDDEN_IN.items():
        txt = texts.get(rel)
        if txt is None:
            continue
        for bad in bads:
            if bad in txt:
                problems.append("[定向废止写法] %s 不应出现：%s"
                                "（改用 label_at / label_gap / label_side 沿线相对定位）"
                                % (rel, bad))

    # 2) REQUIRED
    for rel, tokens in REQUIRED.items():
        txt = texts.get(rel)
        if txt is None:
            problems.append("[缺文件] %s" % rel)
            continue
        for t in tokens:
            if t not in txt:
                problems.append("[关键结论缺失] %s 未包含：%s" % (rel, t))

    # 3) COUPLED
    for rel, txt in sorted(texts.items()):
        for need, then, why in COUPLED:
            if all(k in txt for k in need) and not all(k in txt for k in then):
                problems.append("[成对约束] %s：%s" % (rel, why))

    if problems:
        print("❌ 文档-代码一致性检查未通过（%d 项）：" % len(problems))
        for p in problems:
            print("   -", p)
        return 1
    print("✅ 文档-代码一致性检查通过（扫描 %d 个文件）" % len(files))
    return 0


if __name__ == "__main__":
    sys.exit(main())
