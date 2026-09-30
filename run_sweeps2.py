# -*- coding: utf-8 -*-
"""第二批实验（辅助脚本）：社区数 × 社区文档组织方式扫描、分块长度扫描、
向量 vs 混合检索逐题对比、增量更新计时（重复测量）。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hw4_rag as H  # noqa: E402

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
P = lambda n: os.path.join(OUT, n)
emb = H.Embedder("paraphrase-multilingual-MiniLM-L12-v2")

print("\n########## A) 社区数 × 组织方式扫描（GraphRAG, scope=global）##########")
res = {}
for mode in ("concat", "roundrobin"):
    for c in (1, 2, 3, 4):
        s, _ = H.build_systems(emb, ["graph"], comms=c, global_mode=mode)
        rep = H.evaluate({f"GraphRAG({mode},C={c})": s["GraphRAG"]}, 5, scope="global", quiet=True)
        r = rep["rows"][2]                      # global 类问题
        res[f"{mode}_C{c}"] = rep
        print(f"   {mode:<11} 社区数={c}  全局题 Recall={r['recall']:.3f} MRR={r['mrr']:.3f}")
H.dump_json(P("eval_comms_global.json"), res)

print("\n########## B) 分块长度扫描（Vector RAG）##########")
res = {}
for size in (60, 120, 240):
    s = H.VectorRAG(emb, size=size, overlap=max(8, size // 6))
    rep = H.evaluate({"Vector RAG": s}, 5, quiet=True)
    res[str(size)] = {"overlap": max(8, size // 6), "n_chunks": len(s.chunks), "rows": rep["rows"]}
    print(f"   块长={size} 重叠={max(8, size // 6)} 块数={len(s.chunks)}")
    for r in rep["rows"]:
        print(f"     {r['type']:<7} Recall={r['recall']:.3f} MRR={r['mrr']:.3f}")
H.dump_json(P("eval_chunk.json"), res)

print("\n########## C) Vector vs Hybrid 逐题对比 ##########")
s, _ = H.build_systems(emb, ["vector", "hybrid"])
rows = H.detail_report(s, 5, H.QUESTIONS, "local")
diff = []
for r in rows:
    v, h = r["Vector RAG"], r["Hybrid RAG"]
    if v["first_rank"] != h["first_rank"] or v["miss"] != h["miss"]:
        diff.append({"id": r["id"], "type": r["type"], "src": r["src"], "q": r["q"],
                     "vector": v, "hybrid": h})
        print(f"   {r['id']} {r['type']}: vector rank {v['first_rank']} miss {v['miss']} "
              f"| hybrid rank {h['first_rank']} miss {h['miss']}")
H.dump_json(P("diff_hybrid.json"), diff)

print("\n########## D) 增量更新计时（重复 5 次）##########")
H.dump_json(P("update.json"), H.update_experiment(emb))

print("\n第二批实验完成")
