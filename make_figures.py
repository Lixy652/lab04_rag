# -*- coding: utf-8 -*-
"""
报告配图生成脚本（辅助工具，不是提交物本体）
把实验运行的真实输出渲染成 PNG，供实验报告插图：
  fig1_vector_q01.png   Vector RAG 单题检索明细（含文档号与相似度分数）
  fig2_graph.png        云山大学知识图谱（按社区着色，高亮 Q01 关系链）
  fig3_wiki.png         WikiRAG 条目化知识库样例
  fig4_kscan.png        Top-K 敏感性曲线（三类问题 × 三范式）
  fig5_metrics.png      三范式 × 三类问题的 Recall@5 / MRR 柱状图
  fig6_graph_q11.png    GraphRAG 在全局题 Q11 上的 local / global 对比
"""
import json
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import networkx as nx

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from hw4_rag import CORPUS, ENTITY_ALIASES, PRE_ENTRIES, PRE_TRIPLES  # noqa: E402

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, "out")
IMG = os.path.join(ROOT, "report_img")
os.makedirs(IMG, exist_ok=True)

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DengXian", "SimSun"]
plt.rcParams["axes.unicode_minus"] = False


def read_text(name):
    with open(os.path.join(OUT, name), "r", encoding="utf-8") as f:
        return f.read().replace("\r\n", "\n")


NOISE = ("At line:", "CategoryInfo", "FullyQualifiedErrorId", "python : Warning",
         "Loading weights:", "imits and faster", "+ python", "+  ~", "RemoteException",
         "imits and faster downloads", "Warning: You are sending")


def clean(text):
    """去掉 PowerShell 包装 python 时混入的告警/报错噪声行，只留程序自己的输出；
    并把连续空行压缩为一行（早期导出的日志存在双重换行）。"""
    out = []
    for line in text.split("\n"):
        s = line.strip()
        if any(n in line for n in NOISE):
            continue
        if s.startswith("+") and "~~~" in s:
            continue
        if not s and out and not out[-1].strip():
            continue
        out.append(line.rstrip())
    while out and not out[0].strip():
        out.pop(0)
    return "\n".join(out)


def console_image(text, path, title="", fontsize=11, max_lines=None):
    lines = text.split("\n")
    if max_lines and len(lines) > max_lines:
        lines = lines[:max_lines] + ["... (截断)"]
    n = max(len(lines), 6)
    fig = plt.figure(figsize=(13, 0.26 * n + 1.0), facecolor="#101418")
    fig.text(0.01, 0.985, title, color="#9fe870", fontsize=12, va="top", family="Microsoft YaHei")
    fig.text(0.01, 0.945, "\n".join(lines), color="#e6e6e6", fontsize=fontsize,
             va="top", family="Microsoft YaHei", linespacing=1.45)
    fig.savefig(path, dpi=170, facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)
    print("[fig]", path)


def build_graph():
    G = nx.DiGraph()
    for h, r, t, src in PRE_TRIPLES:
        h, t = ENTITY_ALIASES.get(h, h), ENTITY_ALIASES.get(t, t)
        G.add_edge(h, t, rel=r, src=src)
    comms = nx.community.greedy_modularity_communities(G.to_undirected())
    return G, comms


def community_layout(G, comms):
    """按社区分簇的确定性布局：每个社区一个「环」，度最高的实体放环心，
    社区中心均匀分布在大圆上——避免 45 个中文标签挤成一团。"""
    import math
    order = sorted(range(len(comms)), key=lambda i: -len(comms[i]))
    centers = {}
    R = 2.6
    for k, i in enumerate(order):
        ang = 2 * math.pi * k / max(1, len(order))
        centers[i] = (R * math.cos(ang), R * math.sin(ang))
    pos = {}
    for i, c in enumerate(comms):
        members = list(c)
        hub = max(members, key=lambda n: G.degree(n))
        cx, cy = centers[i]
        ring = [n for n in members if n != hub]
        pos[hub] = (cx, cy)
        r = 0.30 + 0.085 * len(ring)
        for j, n in enumerate(ring):
            ang = 2 * math.pi * j / max(1, len(ring)) + 0.35 * i
            pos[n] = (cx + r * math.cos(ang), cy + r * math.sin(ang))
    return pos


def fig_graph(path):
    G, comms = build_graph()
    colors = ["#4C9BE8", "#F2994A", "#27AE60", "#9B51E0", "#EB5757", "#2D9CDB",
              "#56CCF2", "#F2C94C", "#BB6BD9", "#6FCF97"]
    node_color = {}
    for i, c in enumerate(comms):
        for node in c:
            node_color[node] = colors[i % len(colors)]
    pos = community_layout(G, comms)
    # 拉开仍然贴得太近的节点，避免标签叠在一起
    import math
    for _ in range(400):
        moved = False
        items = list(pos.items())
        for i in range(len(items)):
            for j in range(i + 1, len(items)):
                (n1, p1), (n2, p2) = items[i], items[j]
                dx, dy = p1[0] - p2[0], p1[1] - p2[1]
                dist = math.hypot(dx, dy) or 1e-6
                min_d = 0.12 + 0.020 * (len(n1) + len(n2))
                if dist < min_d:
                    push = (min_d - dist) / 2
                    ux, uy = dx / dist, dy / dist
                    pos[n1] = (p1[0] + ux * push, p1[1] + uy * push)
                    pos[n2] = (p2[0] - ux * push, p2[1] - uy * push)
                    moved = True
        if not moved:
            break
    fig, ax = plt.subplots(figsize=(17, 11.5))
    # 高亮 Q01 关系链：大模型通识课 ←主讲— 李文瀚 —任职于→ 人工智能学院 —院长→ 王启明
    path_edges = [("李文瀚", "大模型通识课"), ("李文瀚", "人工智能学院"), ("人工智能学院", "王启明")]
    normal = [(u, v) for u, v in G.edges() if (u, v) not in path_edges]
    nx.draw_networkx_edges(G, pos, edgelist=normal, ax=ax, edge_color="#b9c2cb",
                           arrows=True, arrowsize=9, width=0.8, alpha=0.75)
    nx.draw_networkx_edges(G, pos, edgelist=path_edges, ax=ax, edge_color="#d62828",
                           arrows=True, arrowsize=18, width=2.6)
    nx.draw_networkx_nodes(G, pos, ax=ax, node_size=1500,
                           node_color=[node_color[n] for n in G.nodes()],
                           edgecolors="white", linewidths=1.2, alpha=0.95)
    nx.draw_networkx_labels(G, pos, ax=ax, font_size=9.5, font_family="Microsoft YaHei",
                            bbox=dict(fc="white", ec="none", alpha=0.72, pad=0.15))
    labels = {(u, v): d["rel"] for u, v, d in G.edges(data=True) if (u, v) in path_edges}
    nx.draw_networkx_edge_labels(G, pos, edge_labels=labels, ax=ax, font_size=10,
                                 font_family="Microsoft YaHei", font_color="#d62828",
                                 bbox=dict(fc="white", ec="none", alpha=0.85))
    handles = [plt.Line2D([], [], marker="o", ls="", color=colors[i % len(colors)],
                          label=f"社区 C{i + 1}（{len(c)} 个实体）", markersize=10)
               for i, c in enumerate(comms)]
    handles.append(plt.Line2D([], [], color="#d62828", lw=2.6, label="Q01 多跳关系链（红色）"))
    ax.legend(handles=handles, loc="upper left", fontsize=11, framealpha=0.9)
    ax.set_title(f"云山大学知识图谱（{G.number_of_nodes()} 个实体 / {G.number_of_edges()} 条三元组 / "
                 f"{len(comms)} 个社区，贪心模块度社区发现）", fontsize=14)
    ax.axis("off")
    ax.set_aspect("equal")
    fig.tight_layout()
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print("[fig]", path)


def fig_entries(path):
    text = read_text("wiki_entries.txt")
    console_image(text, path, title="$ python hw4_rag.py --mode entries   # WikiRAG 条目化知识库",
                  fontsize=10.5, max_lines=60)


def fig_kscan(path, scan):
    fig, axes = plt.subplots(1, 3, figsize=(15, 4.4))
    ks = [1, 2, 3, 5]
    names = [r["paradigm"] for r in scan["1"]["rows"]]
    markers = {"Vector RAG": "o", "GraphRAG": "s", "WikiRAG": "^", "Hybrid RAG": "D"}
    for ax, tp, cn in zip(axes, ("fact", "multi", "global"), ("事实题", "多跳题", "全局题")):
        for name in names:
            ys = [next(r for r in scan[str(k)]["rows"]
                       if r["paradigm"] == name and r["type"] == tp)["recall"] for k in ks]
            ax.plot(ks, ys, marker=markers.get(name, "o"), label=name, linewidth=2)
        ax.set_title(f"{cn} Recall@K", fontsize=13)
        ax.set_xlabel("Top-K（检索窗口）")
        ax.set_ylabel("Recall@K")
        ax.set_xticks(ks)
        ax.set_ylim(-0.03, 1.05)
        ax.grid(alpha=0.3)
        ax.legend(fontsize=9)
    fig.suptitle("检索窗口敏感性：Top-K = 1 / 2 / 3 / 5（内置 15 题）", fontsize=14)
    fig.tight_layout()
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print("[fig]", path)


def fig_metrics(path, ev):
    names = [m["paradigm"] for m in ev["macro"]]
    tps = ["fact", "multi", "global"]
    cn = {"fact": "事实", "multi": "多跳", "global": "全局"}
    x = range(len(tps))
    w = 0.8 / len(names)
    fig, axes = plt.subplots(1, 2, figsize=(14, 4.6))
    for ax, metric, title in zip(axes, ("recall", "mrr"), ("Recall@5", "MRR")):
        for i, name in enumerate(names):
            ys = [next(r for r in ev["rows"] if r["paradigm"] == name and r["type"] == tp)[metric]
                  for tp in tps]
            ax.bar([xi + i * w for xi in x], ys, width=w, label=name)
            for xi, y in zip(x, ys):
                ax.text(xi + i * w, y + 0.02, f"{y:.2f}", ha="center", fontsize=8)
        ax.set_xticks([xi + 0.4 - w / 2 for xi in x])
        ax.set_xticklabels([cn[t] + "题" for t in tps])
        ax.set_ylim(0, 1.12)
        ax.set_title(title, fontsize=13)
        ax.grid(axis="y", alpha=0.3)
    axes[0].legend(fontsize=9)
    fig.suptitle("三范式 × 三类问题的检索质量（20 题，文档级）", fontsize=14)
    fig.tight_layout()
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print("[fig]", path)


def main():
    ev = json.load(open(os.path.join(OUT, "eval.json"), encoding="utf-8"))
    scan = json.load(open(os.path.join(OUT, "scan.json"), encoding="utf-8"))
    console_image(clean(read_text("vector_q01.txt")), os.path.join(IMG, "fig1_vector_q01.png"),
                  title="$ python hw4_rag.py --mode vector --q Q01   # Vector RAG 检索明细（含相似度分数）")
    fig_graph(os.path.join(IMG, "fig2_graph.png"))
    fig_entries(os.path.join(IMG, "fig3_wiki.png"))
    fig_kscan(os.path.join(IMG, "fig4_kscan.png"), scan)
    fig_metrics(os.path.join(IMG, "fig5_metrics.png"), ev)
    both = ("【scope=local：实体锚定 + 沿边走】\n" + clean(read_text("graph_q11_local.txt")) +
            "\n\n【scope=global：社区摘要匹配】\n" + clean(read_text("graph_q11_global.txt")))
    console_image(both, os.path.join(IMG, "fig6_graph_q11.png"),
                  title="$ python hw4_rag.py --mode graph --q Q11 --scope local / global   # 全局题上的两类查询")


if __name__ == "__main__":
    main()
