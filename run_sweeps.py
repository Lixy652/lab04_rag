# -*- coding: utf-8 -*-
"""批量实验驱动（辅助脚本）：一次加载嵌入模型，跑完全部评测/消融/扫描，
避免反复启动进程重复加载模型。所有结果写入 out/*.json。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hw4_rag as H  # noqa: E402

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
os.makedirs(OUT, exist_ok=True)

emb = H.Embedder("paraphrase-multilingual-MiniLM-L12-v2")
P = lambda name: os.path.join(OUT, name)

print("\n########## 1) 主评测：三范式 × 20 题（GraphRAG scope=local）##########")
systems, _ = H.build_systems(emb, ["vector", "graph", "wiki"])
H.dump_json(P("eval.json"), H.evaluate(systems, 5))

print("\n########## 2) GraphRAG scope=global ##########")
graph_only = {k: v for k, v in systems.items() if k == "GraphRAG"}
H.dump_json(P("eval_graph_global.json"), H.evaluate(graph_only, 5, scope="global"))

print("\n########## 3) 社区数消融：--comms 1/2/3/4（graph, local）##########")
for c in (1, 2, 3, 4):
    s, _ = H.build_systems(emb, ["graph"], comms=c)
    rep = H.evaluate(s, 5)
    H.dump_json(P(f"eval_comms{c}.json"), rep)

print("\n########## 4) 实体归一化消融：关掉别名表（graph, local）##########")
s, _ = H.build_systems(emb, ["graph"], normalize=False)
H.dump_json(P("eval_norm_off.json"), H.evaluate(s, 5))
print("-- 归一化开启对照（同配置）--")
s, _ = H.build_systems(emb, ["graph"], normalize=True)
H.dump_json(P("eval_norm_on.json"), H.evaluate(s, 5))

print("\n########## 5) Hybrid RAG（向量 + 字符 bigram，RRF 融合）对比 Vector ##########")
s, _ = H.build_systems(emb, ["vector", "hybrid"])
H.dump_json(P("eval_hybrid.json"), H.evaluate(s, 5))

print("\n########## 6) K 敏感性扫描（K=1/2/3/5，内置 15 题）##########")
builtin = [q for q in H.QUESTIONS if q["src"] == "内置"]
scan = {}
for k in (1, 2, 3, 5):
    s, _ = H.build_systems(emb, ["vector", "graph", "wiki"])
    scan[str(k)] = H.evaluate(s, k, builtin, quiet=True)
    print(f"  K={k} 完成")
names = list(dict.fromkeys(r["paradigm"] for r in scan["1"]["rows"]))
print(f"{'范式':<12}{'类型':<8}" + "".join(f"{'K=' + str(k):>9}" for k in (1, 2, 3, 5)))
for name in names:
    for tp in ("fact", "multi", "global"):
        line = f"{name:<14}{H.TYPE_CN[tp]:<10}"
        for k in (1, 2, 3, 5):
            row = next(r for r in scan[str(k)]["rows"] if r["paradigm"] == name and r["type"] == tp)
            line += f"{row['recall']:>9.3f}"
        print(line)
H.dump_json(P("scan.json"), scan)

print("\n########## 7) 逐题明细（local / global）##########")
H.dump_json(P("detail.json"), {"k": 5, "scope": "local",
                               "rows": H.detail_report(systems, 5, H.QUESTIONS, "local")})
H.dump_json(P("detail_graph_global.json"), {"k": 5, "scope": "global",
                                            "rows": H.detail_report(systems, 5, H.QUESTIONS, "global")})

print("\n########## 8) 增量更新实验：新增 doc16 ##########")
H.dump_json(P("update.json"), H.update_experiment(emb))

print("\n########## 9) 含 doc16 的评测（考察新增文档是否影响存量问题）##########")
s, _ = H.build_systems(emb, ["vector", "graph", "wiki"], doc16=True)
H.dump_json(P("eval_doc16.json"), H.evaluate(s, 5))

print("\n全部实验完成")
