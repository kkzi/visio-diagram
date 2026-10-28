# -*- coding: utf-8 -*-
"""
最小可运行示例：画一张「三环节 + 闭环」的流程图，导出三件套并自检。

跑法（在 skill 目录下）：
    python assets/example_flowchart.py            # 输出到当前目录 out/
    python assets/example_flowchart.py D:/tmp/x   # 指定输出前缀

照抄这个骨架改内容即可 —— 坐标全部是「页面左上角原点 / Y 向下 / mm」。

★ 进程安全：绘图中途抛异常若忘了 close() 会残留 VISIO.EXE（本 skill 的第一大坑）。
  两种写法二选一：
    · try/finally 包住全部绘图，finally 里 d.close()；
    · 或直接 `with VDraw(...) as d:` —— 退出（含异常）自动 close()。
"""

import os
import sys

# ★ Windows 控制台默认 GBK，中文输出虽可过，但统一切 UTF-8 更稳。
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                os.pardir, "scripts"))
from visio_lib import VDraw          # noqa: E402


def main(base="out/示例图_流程"):
    W, H = 160.0, 82.0                        # 画布：按最终印刷宽度 1:1 设计
    d = VDraw(W, H, theme="tech-cn")
    try:
        # ---- 分区（泳道）----
        d.lane(0, 0, W, 58, "")                                # label 留空 = 只画底框
        d.text(0, 1.5, W, 6, "处理链路", size=8, color="gray", align=0)

        # ---- 三个环节 ----
        bw, bh, gap = 42.0, 26.0, 11.0
        x0, y0 = 6.0, 12.0
        steps = [
            ("① 配置与构建", "导入源配置 → 归一为内部模型库\n构建工程并下发"),
            ("② 生成与发送", "按绝对时基调度 → 编码\n→ 组包 → 发送"),
            ("③ 接收与记录", "被测对象解析存储\n输出文本 + 索引文件"),
        ]
        for i, (t, b) in enumerate(steps):
            x = x0 + i * (bw + gap)
            # 第三个框用灰底表示「外部对象」
            fill = "muted_fill" if i == 2 else "white"
            d.step_box(x, y0, bw, bh, t, b, fill=fill)
            if i:
                d.h_arrow(x - gap, y0 + bh / 2, x, color="black")

        # ---- 闭环迭代（虚线回环）----
        yloop = y0 + bh + 10
        d.poly([(x0 + 2 * (bw + gap) + bw / 2, y0 + bh),
                (x0 + 2 * (bw + gap) + bw / 2, yloop),
                (x0 + bw / 2, yloop),
                (x0 + bw / 2, y0 + bh)],
               arrow=4, lpat=4, color="gray", lw="0.9 pt")
        d.text(x0, yloop - 6.5, 3 * bw + 2 * gap, 5.5,
               "闭环迭代（据比对结果调整配置）", size=7, color="gray", align=1)

        # ---- 图例 ----
        d.legend(6, H - 9, W - 12, [("主数据流", None), ("闭环迭代", 4)])

        # ---- 自检 + 出图 ----（先自检后出图，与 SKILL.md 流程一致）
        d.check()
        print("导出：")
        for p in d.export_all(base):
            print("   ", p)
    finally:
        d.close()                             # ★ 异常也要收尾，否则残留 VISIO.EXE


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "out/示例图_流程")
