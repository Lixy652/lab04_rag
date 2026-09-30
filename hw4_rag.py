# -*- coding: utf-8 -*-
"""
实验作业四：Vector RAG / GraphRAG / WikiRAG 对比实验（含本人修改）
对应教材第 8 章《检索增强生成》

本文件以实验指南第七节参考框架为底稿，以下是本人的修改点（代码中以 [修改] 标注）：

  [修改1] 新增 5 道自拟题 Q16~Q20（事实/多跳/全局三类均有覆盖），全部标注 evidence，
          与内置题考察点岔开（见 QUESTIONS 中 "src": "自拟"）。
  [修改2] 新增实体归一化：ENTITY_ALIASES + normalize_text()，修复「云大」↔「云山大学」、
          「《大模型通识课》」↔「大模型通识课」等写法匹配不到图节点的问题（指南 Q4）。
  [修改3] GraphRAG 沿边展开改为「已展开集合 + 出度优先 + 邻域上限」，
          避免在小图上反复回到同一实体、并让链式关系先被展开。
  [修改4] 新增 HybridRAG：向量分与字符 bigram 词面分用 RRF（教材式 8-6）融合（任务五 a）。
  [修改5] 新增 doc16 增量更新实验（--doc16）与三种范式的重建耗时统计（任务五 b）。
  [修改6] 新增 --mode scan（K=1/2/3/5 敏感性扫描）、--mode detail（20 题 × 3 范式逐题明细）、
          --mode faith（LLM 生成 + 忠实度评分素材）、--mode update（构建/更新工程账）。
  [修改7] 评测结果统一导出 UTF-8 JSON（--out），报告中的表格数据全部由脚本输出直接抄录。

常用命令：
  python hw4_rag.py --mode eval                       # 三范式 × 20 题检索评测
  python hw4_rag.py --mode scan --out out/scan.json   # K=1/2/3/5 敏感性扫描
  python hw4_rag.py --mode vector --q Q01             # 单题：检索明细（含分数）+ 生成
  python hw4_rag.py --mode graph  --q Q01 --scope local
  python hw4_rag.py --mode graph  --q Q11 --scope global
  python hw4_rag.py --mode detail --out out/detail.json
  python hw4_rag.py --mode faith  --out out/faith.json
  python hw4_rag.py --mode extract                    # 调 LLM 现场抽三元组（需 API）
  python hw4_rag.py --mode update                     # 新增 doc16 的增量更新账
  python hw4_rag.py --mrr_check 3,1,2                 # MRR 手算复核（式 8-10）
"""
import argparse
import json
import os
import re
import sys
import time
import zlib

# 控制台统一按 UTF-8 输出，避免 Windows 终端乱码
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))

# ============================================================
# 1. 语料库「云山大学」：15 篇虚构文档（三类问题的答案均已埋入）
# ============================================================
CORPUS = {
    "doc01": "云山大学位于岭南省云州市，建于 1958 年，是省属重点大学，现有全日制本科生约 2.1 万人。校训为「格物致知」。",
    "doc02": "人工智能学院成立于 2018 年，首任及现任院长为王启明教授。学院下设机器学习系、智能科学系、认知计算系三个系。",
    "doc03": "计算机学院前身为 1985 年成立的计算机系，现任院长陈国峰教授，设计算机科学与技术、软件工程两个本科专业。",
    "doc04": "人工智能学院教师：李文瀚，教授，2019 年入职，研究方向为检索增强生成，主讲《大模型通识课》；赵婉晴，副教授，研究方向为知识图谱，主讲《知识图谱导论》；孙浩然，讲师，研究方向为强化学习。",
    "doc05": "计算机学院教师：周天宇，教授，研究方向为分布式系统；林小雨，副教授，研究方向为数据库系统。",
    "doc06": "《大模型通识课》课程代码 AI2101，3 学分、32 学时，春季学期开设，授课教师李文瀚，面向全校本科生，无先修课程要求。",
    "doc07": "《知识图谱导论》课程代码 AI3305，2 学分，秋季学期开设，授课教师赵婉晴，面向人工智能学院研究生。",
    "doc08": "选课规则：本科生每学期最多修 30 学分；GPA 低于 2.0 给予学术警告。先修要求：《知识图谱导论》需先修《数据结构》，《大模型通识课》无先修要求。",
    "doc09": "奖学金体系：国家奖学金 8000 元/年，评定比例约 2%；校长奖学金 20000 元/年，全校每年 10 人；云山一等奖学金 3000 元/年，比例约 5%。",
    "doc10": "科研平台：认知计算全国重点实验室依托人工智能学院建设；岭南超算中心由云山大学与省科技厅共建，计算机学院参与运行管理。",
    "doc11": "学生社团：AI 协会，指导教师李文瀚，每周三晚组织论文研读；机器人战队，指导教师孙浩然，每年参加全国机器人大赛。",
    "doc12": "校历：春季学期 3 月 2 日开学、7 月 5 日放暑假；秋季学期 9 月 1 日开学、次年 1 月 15 日放寒假。",
    "doc13": "图书馆藏书 380 万册，开放时间为每天 7:00-22:00；人工智能分馆位于理科楼 B 座 3 层，收藏大模型与智能体专题图书。",
    "doc14": "校园交通：地铁 3 号线「云大站」距东门 200 米；校内校巴共 5 条线路，10 分钟一班。",
    "doc15": "国际交流：学校与 12 个国家的 47 所高校签有交换协议；人工智能学院与新加坡南洋理工大学有本科联合培养项目。",
}

# [修改5] 增量更新实验用：新增第 16 篇文档（默认不启用，--doc16 打开）
DOC16 = {
    "doc16": "信息化建设：云山大学信息化办公室 2024 年上线「云山智课」教学平台，支持课程知识库问答，"
             "由人工智能学院李文瀚团队提供技术支持；平台已接入图书馆人工智能分馆的 1.2 万册电子图书。",
}

# ============================================================
# 2. 问题集：3 类 × 5 题（内置）+ 5 道自拟题；evidence = 支持答案的文档（金标准）
#    fact=局部事实 multi=多跳关系 global=全局归纳
# ============================================================
QUESTIONS = [
    # ---- 内置 15 题 ----
    {"id": "Q01", "type": "multi", "src": "内置", "q": "《大模型通识课》授课教师所在学院的行政负责人是谁？",
     "evidence": ["doc04", "doc06", "doc02"]},
    {"id": "Q02", "type": "multi", "src": "内置", "q": "AI 协会的指导教师主讲课程的课程代码是什么？",
     "evidence": ["doc11", "doc04", "doc06"]},
    {"id": "Q03", "type": "multi", "src": "内置", "q": "认知计算全国重点实验室依托的学院成立于哪一年？",
     "evidence": ["doc10", "doc02"]},
    {"id": "Q04", "type": "multi", "src": "内置", "q": "《知识图谱导论》授课教师的研究方向是什么？",
     "evidence": ["doc07", "doc04"]},
    {"id": "Q05", "type": "multi", "src": "内置", "q": "机器人战队的指导教师的研究方向是什么？",
     "evidence": ["doc11", "doc04"]},
    {"id": "Q06", "type": "fact", "src": "内置", "q": "云山大学的校训是什么？", "evidence": ["doc01"]},
    {"id": "Q07", "type": "fact", "src": "内置", "q": "国家奖学金的金额是多少？", "evidence": ["doc09"]},
    {"id": "Q08", "type": "fact", "src": "内置", "q": "图书馆每天几点到几点开放？", "evidence": ["doc13"]},
    {"id": "Q09", "type": "fact", "src": "内置", "q": "人工智能学院成立于哪一年？", "evidence": ["doc02"]},
    {"id": "Q10", "type": "fact", "src": "内置", "q": "春季学期什么时候开学？", "evidence": ["doc12"]},
    {"id": "Q11", "type": "global", "src": "内置", "q": "云山大学有哪些科研平台？分别依托谁建设？",
     "evidence": ["doc10", "doc02", "doc03"]},
    {"id": "Q12", "type": "global", "src": "内置", "q": "学校的学生奖励体系包含哪些项目？",
     "evidence": ["doc09"]},
    {"id": "Q13", "type": "global", "src": "内置", "q": "人工智能学院有哪些教师？各自做什么研究方向？",
     "evidence": ["doc04"]},
    {"id": "Q14", "type": "global", "src": "内置", "q": "本科生选课有哪些限制和先修要求？",
     "evidence": ["doc08"]},
    {"id": "Q15", "type": "global", "src": "内置", "q": "对想深入学习大模型的学生，学校提供哪些课程和课外活动？",
     "evidence": ["doc06", "doc11"]},

    # ---- [修改1] 自拟 5 题：三类各至少 1 道，考察点与内置题错开 ----
    {"id": "Q16", "type": "fact", "src": "自拟", "q": "人工智能分馆收藏哪一类专题图书？",
     "evidence": ["doc13"]},
    {"id": "Q17", "type": "fact", "src": "自拟", "q": "校长奖学金每年全校有多少人可以拿到？",
     "evidence": ["doc09"]},
    {"id": "Q18", "type": "multi", "src": "自拟", "q": "《知识图谱导论》授课教师所在学院一共下设哪几个系？",
     "evidence": ["doc07", "doc04", "doc02"]},
    {"id": "Q19", "type": "multi", "src": "自拟", "q": "岭南超算中心由哪一方与云山大学共建？参与其运行管理的学院院长是谁？",
     "evidence": ["doc10", "doc03"]},
    {"id": "Q20", "type": "global", "src": "自拟", "q": "云山大学的地理位置、交通与校历安排是怎样的？",
     "evidence": ["doc01", "doc12", "doc14"]},
]


# ============================================================
# 3. 分块（教材 8.3 节：块长≈300 token，中文按 120 字近似，重叠 20 字）
# ============================================================
def chunk_docs(size=120, overlap=20, corpus=None):
    """按固定窗口切块；step = size - overlap；末块不足 size 时一并保留。"""
    chunks, step = [], size - overlap
    for doc_id, text in (corpus or CORPUS).items():
        i = 0
        while True:
            chunks.append((doc_id, text[i:i + size]))
            if i + size >= len(text):
                break
            i += step
    return chunks


def _dedup(scored, k):
    """把「块/三元组邻域/条目」级别的打分结果按文档去重，取 Top-K 文档。"""
    docs, seen = [], set()
    for doc, s in scored:          # scored 已按分数降序
        if doc not in seen:
            docs.append(doc)
            seen.add(doc)
            if len(docs) >= k:
                break
    return docs


def _bigrams(text):
    """[修改4] 字符 bigram 集合，用于 Hybrid RAG 的词面匹配分。"""
    t = re.sub(r"[\s《》「」（）()，。；：、,.;:]", "", text)
    if len(t) < 2:
        return {t} if t else set()
    return {t[i:i + 2] for i in range(len(t) - 1)}


# ============================================================
# 4. 嵌入后端：优先 sentence-transformers（教材 8.4 节）；
#    未安装/无网络时自动降级为字符 3-gram 相似度（流程可跑通，精度有限）
# ============================================================
class Embedder:
    def __init__(self, model_name):
        self.ok = False
        self.model_name = model_name
        try:
            from sentence_transformers import SentenceTransformer
            self.model = SentenceTransformer(model_name)
            self.ok = True
            print(f"[embed] 已加载 {model_name}")
        except Exception as e:
            print(f"[embed] 加载失败（{type(e).__name__}），降级为字符 3-gram 相似度（精度有限）")

    def encode(self, texts):
        import numpy as np
        if self.ok:
            return self.model.encode(list(texts), normalize_embeddings=True)
        M = np.zeros((len(texts), 512))
        for r, t in enumerate(texts):
            t = re.sub(r"\s", "", t)
            for i in range(max(1, len(t) - 2)):
                g = t[i:i + 3] if len(t) >= 3 else t
                M[r, zlib.crc32(g.encode("utf-8")) % 512] += 1
            n = np.linalg.norm(M[r]) or 1.0
            M[r] /= n
        return M


# ============================================================
# 5. Vector RAG（教材 8.2~8.4 节）：块向量 → 余弦 Top-K → 文档序列
# ============================================================
class VectorRAG:
    name = "Vector RAG"

    def __init__(self, emb, size=120, overlap=20, corpus=None):
        t0 = time.time()
        self.emb = emb
        self.chunks = chunk_docs(size, overlap, corpus)
        self.M = emb.encode([c[1] for c in self.chunks])
        self.build_seconds = time.time() - t0

    def chunk_scores(self, query):
        """返回 (文档号, 块文本, 余弦分) 的全量降序列表，供检索明细打印。"""
        qv = self.emb.encode([query])[0]
        scores = self.M @ qv
        order = scores.argsort()[::-1]
        return [(self.chunks[i][0], self.chunks[i][1], float(scores[i])) for i in order]

    def retrieve(self, query, k=5, scope=None):
        hits = [(d, s) for d, _, s in self.chunk_scores(query)]
        return _dedup(hits, k)


# ============================================================
# 6. Hybrid RAG（任务五 a）：向量分 + 字符 bigram 词面分，RRF 融合（式 8-6）
#    RRF(d) = Σ 1 / (rrf_k + rank_i(d))
# ============================================================
class HybridRAG(VectorRAG):
    name = "Hybrid RAG"

    def retrieve(self, query, k=5, scope=None, rrf_k=60):
        vec = self.chunk_scores(query)
        rank_v = {}
        for r, (d, _t, _s) in enumerate(vec, 1):
            rank_v.setdefault(d, r)
        qb = _bigrams(query)
        lex = []
        for doc_id, text in self.chunks:
            tb = _bigrams(text)
            inter = len(qb & tb)
            lex.append((doc_id, inter / (len(qb) or 1)))
        rank_l = {}
        for r, (d, _s) in enumerate(sorted(lex, key=lambda x: -x[1]), 1):
            rank_l.setdefault(d, r)
        BIG = 10 ** 6
        fused = {d: 1.0 / (rrf_k + rank_v.get(d, BIG)) + 1.0 / (rrf_k + rank_l.get(d, BIG))
                 for d in set(rank_v) | set(rank_l)}
        return [d for d, _ in sorted(fused.items(), key=lambda kv: -kv[1])][:k]


# ============================================================
# 7. GraphRAG（教材 8.7 节简化版）
#    预抽取三元组 (头实体, 关系, 尾实体, 来源文档)；--mode extract 可现场抽取
# ============================================================
PRE_TRIPLES = [
    ("李文瀚", "任职于", "人工智能学院", "doc04"), ("李文瀚", "职称", "教授", "doc04"),
    ("李文瀚", "入职年份", "2019 年", "doc04"), ("李文瀚", "研究方向", "检索增强生成", "doc04"),
    ("李文瀚", "主讲", "大模型通识课", "doc06"), ("李文瀚", "指导", "AI 协会", "doc11"),
    ("赵婉晴", "任职于", "人工智能学院", "doc04"), ("赵婉晴", "研究方向", "知识图谱", "doc04"),
    ("赵婉晴", "主讲", "知识图谱导论", "doc07"), ("孙浩然", "任职于", "人工智能学院", "doc04"),
    ("孙浩然", "研究方向", "强化学习", "doc04"), ("孙浩然", "指导", "机器人战队", "doc11"),
    ("人工智能学院", "院长", "王启明", "doc02"), ("人工智能学院", "成立于", "2018 年", "doc02"),
    ("人工智能学院", "下设系", "机器学习系", "doc02"), ("人工智能学院", "下设系", "智能科学系", "doc02"),
    ("人工智能学院", "下设系", "认知计算系", "doc02"), ("认知计算全国重点实验室", "依托", "人工智能学院", "doc10"),
    ("岭南超算中心", "参与共建", "计算机学院", "doc10"), ("计算机学院", "院长", "陈国峰", "doc03"),
    ("计算机学院", "前身", "计算机系", "doc03"), ("周天宇", "任职于", "计算机学院", "doc05"),
    ("周天宇", "研究方向", "分布式系统", "doc05"), ("林小雨", "任职于", "计算机学院", "doc05"),
    ("林小雨", "研究方向", "数据库系统", "doc05"), ("大模型通识课", "课程代码", "AI2101", "doc06"),
    ("大模型通识课", "学分", "3 学分", "doc06"), ("大模型通识课", "开设学期", "春季学期", "doc06"),
    ("知识图谱导论", "课程代码", "AI3305", "doc07"), ("知识图谱导论", "先修课程", "数据结构", "doc08"),
    ("AI 协会", "指导教师", "李文瀚", "doc11"), ("机器人战队", "指导教师", "孙浩然", "doc11"),
    ("国家奖学金", "金额", "8000 元/年", "doc09"), ("校长奖学金", "金额", "20000 元/年", "doc09"),
    ("云山一等奖学金", "金额", "3000 元/年", "doc09"), ("云山大学", "校训", "格物致知", "doc01"),
    ("云山大学", "建校于", "1958 年", "doc01"), ("云山大学", "位于", "岭南省云州市", "doc01"),
    ("图书馆", "藏书量", "380 万册", "doc13"), ("春季学期", "开学日期", "3 月 2 日", "doc12"),
]

# [修改5] doc16 带来的增量三元组（仅 --doc16 时并入）
DOC16_TRIPLES = [
    ("云山智课平台", "上线单位", "信息化办公室", "doc16"),
    ("云山智课平台", "技术支持", "李文瀚", "doc16"),
    ("云山智课平台", "接入", "人工智能分馆", "doc16"),
]

# 社区摘要（教材 8.7 节由 LLM 生成；此处预置 3 条，对应图的三个稠密子区）
COMMUNITY_SUMMARIES = [
    ("社区 C1 教学与课程：人工智能学院设三个系，开设《大模型通识课》（AI2101，李文瀚主讲）与"
     "《知识图谱导论》（AI3305，赵婉晴主讲），AI3305 需先修《数据结构》。",
     ["doc02", "doc04", "doc06", "doc07", "doc08"]),
    ("社区 C2 科研与国际：认知计算全国重点实验室依托人工智能学院，岭南超算中心由计算机学院参与共建；"
     "学院与新加坡南洋理工大学有联合培养项目。",
     ["doc02", "doc03", "doc10", "doc15"]),
    ("社区 C3 学生生活：奖学金分国家（8000 元）、校长（20000 元）、云山一等（3000 元）三级；"
     "社团有 AI 协会与机器人战队，图书馆藏书 380 万册、7:00-22:00 开放。",
     ["doc09", "doc11", "doc13"]),
]

# [修改2] 实体归一化表：别名 / 带书名号写法 → 规范实体名（指南 Q4：「云大」与「云山大学」合并）
ENTITY_ALIASES = {
    "云大": "云山大学", "云山大学": "云山大学", "云大站": "云山大学",
    "人工智能学院": "人工智能学院", "AI 学院": "人工智能学院", "智能学院": "人工智能学院",
    "计算机学院": "计算机学院", "计院": "计算机学院",
    "大模型通识课": "大模型通识课",
    "知识图谱导论": "知识图谱导论",
    "AI 协会": "AI 协会", "AI协会": "AI 协会",
    "认知计算全国重点实验室": "认知计算全国重点实验室", "认知计算实验室": "认知计算全国重点实验室",
    "岭南超算中心": "岭南超算中心", "超算中心": "岭南超算中心",
    "机器人战队": "机器人战队",
}


def normalize_text(s):
    """去掉书名号、引号与空白，便于实体匹配（不改动图里的规范名）。"""
    return re.sub(r"[\s《》「」【】（）()]", "", s or "")


def match_entities(query, nodes, topn=2, use_alias=True):
    """[修改2] 先按别名表命中规范实体，再退回「节点名子串出现在问题里」的匹配。
    use_alias=False 时退化为参考实现的原始匹配方式（用于归一化消融实验）。"""
    nq = normalize_text(query) if use_alias else query
    hits = []
    for alias, canon in (ENTITY_ALIASES.items() if use_alias else []):
        if normalize_text(alias) in nq and canon in nodes:
            hits.append(canon)
    for n in nodes:
        key = normalize_text(n) if use_alias else n
        if key in nq:
            hits.append(n)
    out, seen = [], set()
    for e in sorted(set(hits), key=len, reverse=True):
        if e not in seen:
            seen.add(e)
            out.append(e)
    return out[:topn]


class GraphRAG:
    name = "GraphRAG"

    def __init__(self, emb, triples=None, hops=3, top_communities=2, normalize=True,
                 global_mode="concat"):
        import networkx as nx
        t0 = time.time()
        self.emb = emb
        self.hops = hops
        self.top_communities = top_communities
        self.normalize = normalize
        self.global_mode = global_mode
        self.triples = list(triples if triples is not None else PRE_TRIPLES)
        self.G = nx.DiGraph()
        for h, r, t, src in self.triples:
            # [修改2] 建图时对头尾实体做归一，保证同一实体只有一种写法
            if normalize:
                h, t = ENTITY_ALIASES.get(h, h), ENTITY_ALIASES.get(t, t)
            self.G.add_edge(h, t, rel=r, src=src)
        try:  # 教材 8.7 节用 Louvain；此处用 networkx 内置贪心模块度社区发现
            self.comms = nx.community.greedy_modularity_communities(self.G.to_undirected())
        except Exception:
            self.comms = [set(self.G.nodes)]
        self.summ_M = emb.encode([s for s, _ in COMMUNITY_SUMMARIES])
        self.build_seconds = time.time() - t0
        print(f"[graph] 实体 {self.G.number_of_nodes()} 个 / 三元组 {len(self.triples)} 条 / 社区 {len(self.comms)} 个")

    def entities(self, query, topn=2):
        return match_entities(query, set(self.G.nodes), topn, use_alias=self.normalize)

    def retrieve(self, query, k=5, scope="local"):
        return self._local(query, k) if scope == "local" else self._global(query, k)

    def walk(self, query, hops=None, frontier_cap=8):
        """[修改3] 沿边走：返回 (命中的三元组文本, 来源文档) 列表，供明细打印。"""
        hops = hops or self.hops
        ents = self.entities(query)
        seen_edges, scored = set(), []
        expanded = set()
        frontier = list(ents)
        trace = [f"起始实体：{ents or '（未匹配到实体）'}"]
        for h in range(hops):
            nxt = []
            for e in frontier:
                if e in expanded:
                    continue
                expanded.add(e)
                edges = list(self.G.out_edges(e, data=True)) + list(self.G.in_edges(e, data=True))
                for hh, tt, d in edges:
                    key = (hh, d["rel"], tt)
                    if key in seen_edges:
                        continue
                    seen_edges.add(key)
                    scored.append((d["src"], f"{hh} —{d['rel']}→ {tt}"))
                    other = tt if hh == e else hh
                    if other not in expanded:
                        nxt.append(other)
            # 出度>0 的实体优先展开，保证关系链先走通
            nxt.sort(key=lambda x: 0 if self.G.out_degree(x) > 0 else 1)
            frontier = nxt[:frontier_cap]
            trace.append(f"第 {h + 1} 跳：展开 {len(expanded)} 个实体，累计三元组 {len(scored)} 条")
            if not frontier:
                break
        return scored, trace

    def _local(self, query, k):
        scored, _trace = self.walk(query)
        if not scored:                      # 未命中实体 → 退化为全局查询
            return self._global(query, k)
        qv = self.emb.encode([query])[0]
        sims = self.emb.encode([t for _, t in scored]) @ qv
        best = {}
        for (doc, txt), s in zip(scored, sims):
            if doc not in best or s > best[doc]:
                best[doc] = float(s)
        ranked = sorted(best.items(), key=lambda kv: -kv[1])
        return _dedup(ranked, k)

    def _global(self, query, k):
        qv = self.emb.encode([query])[0]
        scores = self.summ_M @ qv
        top = list(scores.argsort()[::-1][: self.top_communities])   # 取最相关的 N 个社区
        doc_lists = [list(COMMUNITY_SUMMARIES[ci][1]) for ci in top]
        docs = []
        if self.global_mode == "roundrobin":
            # [修改3b] 轮转交错：每个社区轮流放一篇，避免第一个社区把 Top-K 窗口占满
            for pos in range(max(len(x) for x in doc_lists)):
                for lst in doc_lists:
                    if pos < len(lst) and lst[pos] not in docs:
                        docs.append(lst[pos])
        else:                       # 参考实现：按社区顺序拼接后再截断
            for lst in doc_lists:
                for d in lst:
                    if d not in docs:
                        docs.append(d)
        return docs[:k]

    def global_detail(self, query):
        qv = self.emb.encode([query])[0]
        scores = self.summ_M @ qv
        order = scores.argsort()[::-1]
        return [(int(i), float(scores[i]), COMMUNITY_SUMMARIES[i][1]) for i in order]


# ============================================================
# 8. WikiRAG（教材 8.8 节）：条目化知识库，检索单位是「条目」
#    条目 = 标题 + 聚合正文（重写而非复制原文），映射回来源文档评测
# ============================================================
PRE_ENTRIES = [
    ("云山大学", "云山大学位于岭南省云州市，建于 1958 年，省属重点大学，本科生约 2.1 万人，校训「格物致知」。"
     "地铁 3 号线「云大站」距东门 200 米，校巴 5 条线路。校历：春季学期 3 月 2 日开学。",
     ["doc01", "doc12", "doc14"]),
    ("人工智能学院", "人工智能学院成立于 2018 年，院长王启明教授，下设机器学习系、智能科学系、认知计算系。"
     "认知计算全国重点实验室依托本院建设。与新加坡南洋理工大学有本科联合培养项目（每年 15 人）。",
     ["doc02", "doc10", "doc15"]),
    ("计算机学院", "计算机学院前身是 1985 年成立的计算机系，院长陈国峰教授，设计算机科学与技术、软件工程两个专业。"
     "参与岭南超算中心运行管理。",
     ["doc03", "doc10"]),
    ("人工智能学院教师", "李文瀚：教授，2019 年入职，研究检索增强生成，主讲《大模型通识课》，指导 AI 协会。"
     "赵婉晴：副教授，研究知识图谱，主讲《知识图谱导论》。孙浩然：讲师，研究强化学习，指导机器人战队。",
     ["doc04", "doc11"]),
    ("计算机学院教师", "周天宇：教授，研究分布式系统。林小雨：副教授，研究数据库系统。", ["doc05"]),
    ("大模型通识课", "《大模型通识课》（AI2101）：3 学分 32 学时，春季学期开设，李文瀚主讲，面向全校本科生，"
     "无先修要求，教材《人工智能——深度学习大模型智能体》。", ["doc06"]),
    ("知识图谱导论", "《知识图谱导论》（AI3305）：2 学分，秋季学期，赵婉晴主讲，面向人工智能学院研究生，"
     "需先修《数据结构》。", ["doc07", "doc08"]),
    ("奖学金体系", "国家奖学金 8000 元/年（约 2%）；校长奖学金 20000 元/年（全校 10 人）；"
     "云山一等奖学金 3000 元/年（约 5%）。", ["doc09"]),
    ("学生社团与校园生活", "AI 协会：指导教师李文瀚，每周三晚论文研读。机器人战队：指导教师孙浩然，"
     "参加全国机器人大赛。图书馆藏书 380 万册，每天 7:00-22:00 开放，人工智能分馆在理科楼 B 座。",
     ["doc11", "doc13"]),
]

# [修改5] doc16 对应的增量条目
DOC16_ENTRIES = [
    ("云山智课平台", "「云山智课」教学平台由信息化办公室 2024 年上线，支持课程知识库问答，"
     "人工智能学院李文瀚团队提供技术支持，已接入图书馆人工智能分馆 1.2 万册电子图书。", ["doc16"]),
]


class WikiRAG:
    name = "WikiRAG"

    def __init__(self, emb, entries=None, topn_entries=3):
        t0 = time.time()
        self.emb = emb
        self.entries = list(entries if entries is not None else PRE_ENTRIES)
        self.topn_entries = topn_entries
        self.M = emb.encode([b for _, b, _ in self.entries])
        self.build_seconds = time.time() - t0

    def entry_scores(self, query):
        qv = self.emb.encode([query])[0]
        scores = self.M @ qv
        order = scores.argsort()[::-1]
        return [(int(i), float(scores[i])) for i in order]

    def retrieve(self, query, k=5, scope=None):
        order = [i for i, _ in self.entry_scores(query)][: self.topn_entries]   # 条目级 Top-3
        docs = []
        for i in order:
            for d in self.entries[i][2]:
                if d not in docs:
                    docs.append(d)
        return docs[:k]


# ============================================================
# 9. 评测（教材式 8-9 / 8-10）：检索单位（块/邻域三元组/条目）统一映射回文档
# ============================================================
TYPE_CN = {"fact": "事实", "multi": "多跳", "global": "全局"}


def evaluate(systems, k=5, questions=None, quiet=False, scope="local"):
    questions = questions or QUESTIONS
    result = {"k": k, "scope": scope, "rows": [], "systems": list(systems.keys())}
    if not quiet:
        print(f"\n===== 检索评测（文档级，Top-{k}，共 {len(questions)} 题，GraphRAG scope={scope}） =====")
        print(f"{'范式':<12}{'问题类型':<10}{'Recall@' + str(k):>10}{'MRR':>8}  完整命中")
    for name, s in systems.items():
        for tp in ("fact", "multi", "global"):
            qs = [q for q in questions if q["type"] == tp]
            if not qs:
                continue
            recs, rrs, full = [], [], 0
            for q in qs:
                docs = s.retrieve(q["q"], k, scope)
                ev = set(q["evidence"])
                got = ev & set(docs[:k])
                recs.append(len(got) / len(ev))
                rank = next((i + 1 for i, d in enumerate(docs) if d in ev), 0)
                rrs.append(1.0 / rank if rank else 0.0)
                full += (got == ev)
            row = {"paradigm": name, "type": tp, "n": len(qs),
                   "recall": sum(recs) / len(recs), "mrr": sum(rrs) / len(rrs), "full": full}
            result["rows"].append(row)
            if not quiet:
                print(f"{name:<14}{TYPE_CN[tp]:<12}{row['recall']:>8.3f}{row['mrr']:>8.3f}  {full}/{len(qs)}")
    # 三种类型的宏平均
    for name in systems:
        rows = [r for r in result["rows"] if r["paradigm"] == name]
        result.setdefault("macro", []).append({
            "paradigm": name,
            "recall": sum(r["recall"] for r in rows) / len(rows),
            "mrr": sum(r["mrr"] for r in rows) / len(rows),
        })
    if not quiet:
        print("-- 三类宏平均 --")
        for m in result["macro"]:
            print(f"{m['paradigm']:<14}{'':<12}{m['recall']:>8.3f}{m['mrr']:>8.3f}")
        print("\n 提示：把本表抄进报告表 3；另抽 3 题手算 MRR，用 --mrr_check 复核（式 8-10）。")
    return result


# ============================================================
# 10. LLM 生成环节（可选）：任一 OpenAI 兼容接口
#     环境变量：LLM_API_BASE / LLM_API_KEY / LLM_MODEL
# ============================================================
EXTRACT_PROMPT = """从下面这篇大学介绍文档中抽取「头实体-关系-尾实体」三元组，只输出 JSON 数组，
不要输出任何其他文字，格式如 [["李文瀚","任职于","人工智能学院"]]。
关系用 2~4 字短语（任职于/研究方向/主讲/成立于 等），一篇文档抽 5~10 条。
文档：
{doc}"""

EXTRACT_RETRY_PROMPT = """上一次的输出不是合法 JSON。请重新抽取，并且只输出一个 JSON 数组，
不要 Markdown 代码块、不要解释文字。数组元素是三元组字符串数组，例如
[["李文瀚","任职于","人工智能学院"],["李文瀚","研究方向","检索增强生成"]]。
文档：
{doc}"""

ANSWER_PROMPT = """你是问答助手。只依据下面资料回答问题；资料不足以回答时明确说「资料中未提及」，禁止编造。
问题：{q}
资料：
{ctx}
用 1~3 句话回答，句末用括号注明用到的资料编号。"""

# 忠实度评分的对照条件：去掉「禁止编造」约束（并要求不许拒答），
# 用于确认 1~5 分刻度确实有区分度，而不是「怎么问都是 5 分」。
ANSWER_PROMPT_LOOSE = """你是问答助手。请直接回答问题，不要回答「资料中未提及」或「不知道」，
资料不全时可以结合常识补充，务必给出一个完整答案。
问题：{q}
资料：
{ctx}
用 1~3 句话回答。"""


def llm_chat(prompt, temperature=0.0):
    from openai import OpenAI
    base = os.environ.get("LLM_API_BASE")
    key = os.environ.get("LLM_API_KEY")
    model = os.environ.get("LLM_MODEL")
    if not (base and key and model):
        sys.exit("未配置 LLM：请设置环境变量 LLM_API_BASE / LLM_API_KEY / LLM_MODEL")
    client = OpenAI(base_url=base, api_key=key)
    resp = client.chat.completions.create(model=model, temperature=temperature,
                                          messages=[{"role": "user", "content": prompt}])
    return resp.choices[0].message.content


def parse_triples(raw):
    """[人工重点检查处] 从模型输出里截取 JSON 数组；失败抛异常。"""
    try:
        arr = json.loads(raw[raw.index("["): raw.rindex("]") + 1])
    except Exception as e:
        raise ValueError(f"JSON 解析失败：{e}")
    out = []
    for item in arr:
        if isinstance(item, (list, tuple)) and len(item) >= 3:
            out.append([str(item[0]).strip(), str(item[1]).strip(), str(item[2]).strip()])
    if not out:
        raise ValueError("JSON 合法但未解析出任何三元组")
    return out


def extract_triples(out_path=None, limit=None):
    t0, out, bad, retried, ok_retry = time.time(), [], [], 0, 0
    docs = list(CORPUS.items())[: limit or len(CORPUS)]
    detail = []
    for doc_id, text in docs:
        raw = llm_chat(EXTRACT_PROMPT.replace("{doc}", text))
        try:
            arr = parse_triples(raw)
        except Exception:
            retried += 1
            print(f"[warn] {doc_id} 首次解析失败，重试一次：{raw[:60]}...")
            raw2 = llm_chat(EXTRACT_RETRY_PROMPT.replace("{doc}", text))
            try:
                arr = parse_triples(raw2)
                ok_retry += 1
            except Exception:
                bad.append(doc_id)
                print(f"[warn] {doc_id} 重试仍失败，跳过")
                continue
        # [人工检查处] 实体归一化 + 去重
        norm, seen = [], set()
        for h, r, t in arr:
            h2, t2 = ENTITY_ALIASES.get(h, h), ENTITY_ALIASES.get(t, t)
            if (h2, r, t2) in seen:
                continue
            seen.add((h2, r, t2))
            norm.append((h2, r, t2, doc_id))
        out += norm
        detail.append({"doc": doc_id, "raw_n": len(arr), "kept_n": len(norm),
                       "triples": [list(x[:3]) for x in norm]})
        print(f"{doc_id}: 抽到 {len(arr)} 条 / 归一后保留 {len(norm)} 条")
    rep = {"total": len(out), "docs": len(docs), "parse_fail": len(bad),
           "parse_fail_docs": bad, "retried": retried, "retry_ok": ok_retry,
           "seconds": round(time.time() - t0, 1), "detail": detail}
    print(f"\n 共 {len(out)} 条 / 首次解析失败 {retried} 篇（重试成功 {ok_retry} 篇，最终失败 {len(bad)} 篇），"
          f"解析失败率 {len(bad) / max(1, len(docs)):.1%}，耗时 {rep['seconds']}s。")
    print("请人工抽检后，把高质量三元组替换进脚本 PRE_TRIPLES。")
    if out_path:
        dump_json(out_path, {"report": rep, "triples": [list(x) for x in out]})
    return rep


# ============================================================
# 11. 结果落盘与检索明细
# ============================================================
def dump_json(path, obj):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    print(f"[out] 已写入 {path}")


def build_systems(emb, which, doc16=False, size=120, overlap=20, hops=3,
                  comms=2, normalize=True, global_mode="concat"):
    """按 --mode 构造需要的范式；doc16=True 时并入增量文档/三元组/条目。"""
    corpus = dict(CORPUS)
    triples = list(PRE_TRIPLES)
    entries = list(PRE_ENTRIES)
    if doc16:
        corpus.update(DOC16)
        triples += DOC16_TRIPLES
        entries += DOC16_ENTRIES
    systems = {}
    if "vector" in which:
        systems["Vector RAG"] = VectorRAG(emb, size, overlap, corpus)
    if "hybrid" in which:
        systems["Hybrid RAG"] = HybridRAG(emb, size, overlap, corpus)
    if "graph" in which:
        systems["GraphRAG"] = GraphRAG(emb, triples, hops, comms, normalize, global_mode)
    if "wiki" in which:
        systems["WikiRAG"] = WikiRAG(emb, entries)
    return systems, corpus


def detail_report(systems, k=5, questions=None, scope="local"):
    """[修改6] 20 题 × 各范式逐题明细：检索到的文档、首次命中排名、是否漏检。"""
    questions = questions or QUESTIONS
    rows = []
    for q in questions:
        row = {"id": q["id"], "type": q["type"], "src": q["src"], "q": q["q"], "evidence": q["evidence"]}
        for name, s in systems.items():
            docs = s.retrieve(q["q"], k, scope)
            ev = set(q["evidence"])
            rank = next((i + 1 for i, d in enumerate(docs) if d in ev), 0)
            row[name] = {"docs": docs, "first_rank": rank,
                         "hit": sorted(ev & set(docs[:k])), "miss": sorted(ev - set(docs[:k]))}
        rows.append(row)
    return rows


def print_single(systems, q, k, scope, emb_ok):
    """单题：打印检索明细（含分数）、映射回的文档、以及 LLM 生成（若配置）。"""
    print(f"\n[{q['id']}·{TYPE_CN[q['type']]}·{q['src']}] {q['q']}")
    print(f"金标准 evidence：{q['evidence']}")
    for name, s in systems.items():
        print(f"\n--- {name} ---")
        if isinstance(s, HybridRAG):
            print("  向量分与字符 bigram 词面分经 RRF 融合（式 8-6），融合后文档序列：")
            docs = s.retrieve(q["q"], k, scope)
            print(f"    {docs}")
        elif isinstance(s, VectorRAG):
            hits = s.chunk_scores(q["q"])[:8]
            print("  块级 Top-8（文档号 | 余弦分 | 块文本前 40 字）：")
            for d, t, sc in hits:
                print(f"    {d}  {sc:+.4f}  {t[:40]}")
            docs = _dedup([(d, sc) for d, _t, sc in s.chunk_scores(q["q"])], k)
        elif isinstance(s, GraphRAG) and scope == "local":
            scored, trace = s.walk(q["q"])
            print("  实体链接与沿边展开：")
            print("    " + "\n    ".join(trace))
            print("  命中的三元组（去重后前 12 条）：")
            for doc, txt in scored[:12]:
                print(f"    [{doc}] {txt}")
            docs = s.retrieve(q["q"], k, scope)
        elif isinstance(s, GraphRAG):
            print("  社区摘要相似度：")
            for i, sc, srcs in s.global_detail(q["q"]):
                print(f"    C{i + 1}  {sc:+.4f}  来源文档 {srcs}")
            docs = s.retrieve(q["q"], k, scope)
        else:
            print("  条目级 Top-3：")
            for i, sc in s.entry_scores(q["q"])[:3]:
                print(f"    {s.entries[i][0]}  {sc:+.4f}  来源文档 {s.entries[i][2]}")
            docs = s.retrieve(q["q"], k, scope)
        rank = next((i + 1 for i, d in enumerate(docs) if d in set(q["evidence"])), 0)
        got = sorted(set(q["evidence"]) & set(docs))
        print(f"  → 映射回文档（Top-{k}）：{docs}")
        print(f"  → 命中 {got} / 金标准 {q['evidence']}，首次命中排名 {rank or '未命中'}")
    primary = systems[list(systems)[0]]
    docs = primary.retrieve(q["q"], k, scope)
    ctx = "\n".join(f"[{d}] {CORPUS.get(d, DOC16.get(d, ''))}" for d in docs)
    try:
        print("\n【LLM 回答】", llm_chat(ANSWER_PROMPT.replace("{q}", q["q"]).replace("{ctx}", ctx)))
    except SystemExit:
        print("\n（未配置 LLM API，仅输出检索结果与上下文）\n" + ctx)


def faith_report(systems, qids, out_path, loose=False):
    """[修改6] 忠实度评分素材：对指定题目，逐范式检索 → LLM 生成 → 落地 JSON。
    loose=True 时使用不带「禁止编造」约束的对照 prompt，用于展示幻觉。"""
    rows = []
    prompt = ANSWER_PROMPT_LOOSE if loose else ANSWER_PROMPT
    for qid in qids:
        q = next(x for x in QUESTIONS if x["id"] == qid)
        for name, s in systems.items():
            docs = s.retrieve(q["q"], 5)
            ctx = "\n".join(f"[{d}] {CORPUS[d]}" for d in docs)
            try:
                ans = llm_chat(prompt.replace("{q}", q["q"]).replace("{ctx}", ctx))
            except SystemExit:
                ans = "（未配置 LLM API）"
            rows.append({"id": qid, "type": q["type"], "paradigm": name, "q": q["q"],
                         "evidence": q["evidence"], "docs": docs, "ctx": ctx,
                         "prompt": "loose" if loose else "strict", "answer": ans})
            print(f"[faith] {qid} {name} 检索 {docs}\n        回答：{ans}\n")
    dump_json(out_path, {"rows": rows})
    return rows


def print_entries():
    print(f"\n===== WikiRAG 条目库（{len(PRE_ENTRIES)} 条）=====")
    tot = 0
    for title, body, srcs in PRE_ENTRIES:
        tot += len(body)
        print(f"\n[{title}]  正文 {len(body)} 字  来源 {srcs}\n  {body}")
    print(f"\n平均条目长度：{tot / len(PRE_ENTRIES):.1f} 字")


def print_triples():
    print(f"\n===== PRE_TRIPLES（{len(PRE_TRIPLES)} 条，前 12 条样例）=====")
    for h, r, t, src in PRE_TRIPLES[:12]:
        print(f"  ({h}, {r}, {t})  来源 {src}")
    print("\n===== 实体归一化前后对照（重叠出现的同一实体）=====")
    raw = {}
    for h, r, t, src in PRE_TRIPLES:
        for e in (h, t):
            raw.setdefault(ENTITY_ALIASES.get(e, e), set()).add(e)
    for canon, forms in raw.items():
        if len(forms) > 1:
            print(f"  {sorted(forms)} → {canon}")


def update_experiment(emb, k=5, repeats=5):
    """[修改5] 增量更新实验：新增 doc16 前后，三种范式各要重建什么、耗时多少。
    构建耗时重复 repeats 次取均值，避免单次测量被调度噪声污染。"""
    print("\n===== 增量更新实验：在 CORPUS 中新增 doc16 =====")
    base_s, new_s = {}, {}
    t_base = t_new = 0.0
    for name in ("Vector RAG", "GraphRAG", "WikiRAG"):
        key = {"Vector RAG": "vector", "GraphRAG": "graph", "WikiRAG": "wiki"}[name]
        tb, tn = [], []
        for _ in range(repeats):
            t0 = time.time()
            b, _ = build_systems(emb, [key])
            tb.append(time.time() - t0)
            t0 = time.time()
            n, _ = build_systems(emb, [key], doc16=True)
            tn.append(time.time() - t0)
            base_s[name], new_s[name] = b[name], n[name]
        t_base, t_new = sum(tb) / len(tb), sum(tn) / len(tn)
        base_s[name].build_seconds = t_base
        new_s[name].build_seconds = t_new
        base_s[name].timing = (min(tb), max(tb))
        new_s[name].timing = (min(tn), max(tn))
    print(f"\n{'范式':<12}{'全量重建(15 篇)':>18}{'含 doc16 重建':>16}{'增量':>10}   新增改动")
    for name in base_s:
        b, n = base_s[name].build_seconds, new_s[name].build_seconds
        what = {"Vector RAG": "新增 1 个块并重新编码（+1 条向量）",
                "GraphRAG": "新增 3 条三元组 / 3 个实体，重建图与社区",
                "WikiRAG": "新增 1 个条目并重新编码（+1 条向量）"}[name]
        print(f"{name:<14}{b:>15.3f}s{n:>15.3f}s{n - b:>9.3f}s   {what}")
    g0, g1 = base_s["GraphRAG"], new_s["GraphRAG"]
    print(f"\n图规模：{g0.G.number_of_nodes()} 实体 / {len(g0.triples)} 三元组 / {len(g0.comms)} 社区 "
          f"→ {g1.G.number_of_nodes()} 实体 / {len(g1.triples)} 三元组 / {len(g1.comms)} 社区")
    print("注：嵌入模型加载（约 5~10s）与 LLM 抽取/条目改写成本不计入上表，见报告 4.4 节。")
    return {"base_seconds": {n: base_s[n].build_seconds for n in base_s},
            "with_doc16_seconds": {n: new_s[n].build_seconds for n in new_s},
            "timing_range": {n: {"base": base_s[n].timing, "doc16": new_s[n].timing} for n in base_s},
            "graph_before": {"nodes": g0.G.number_of_nodes(), "triples": len(g0.triples), "comms": len(g0.comms)},
            "graph_after": {"nodes": g1.G.number_of_nodes(), "triples": len(g1.triples), "comms": len(g1.comms)}}


# ============================================================
# 12. 主入口
# ============================================================
def main():
    ap = argparse.ArgumentParser(description="实验作业四：三种 RAG 对比实验")
    ap.add_argument("--mode", default="eval",
                    choices=["eval", "vector", "graph", "wiki", "hybrid", "extract",
                             "scan", "detail", "faith", "entries", "triples", "update"])
    ap.add_argument("--q", default="Q01", help="问题编号，如 Q01；faith/detail 可用逗号分隔")
    ap.add_argument("--scope", choices=["local", "global"], default="local")
    ap.add_argument("--k", type=int, default=5)
    ap.add_argument("--emb", default="paraphrase-multilingual-MiniLM-L12-v2")
    ap.add_argument("--mrr_check", default=None, help="如 3,1,2：打印 MRR 手算过程")
    ap.add_argument("--doc16", action="store_true", help="并入 doc16（增量更新实验）")
    ap.add_argument("--hops", type=int, default=3, help="GraphRAG 沿边最大跳数")
    ap.add_argument("--comms", type=int, default=2, help="GraphRAG 全局查询取前 N 个社区（参考实现为 2）")
    ap.add_argument("--no_norm", action="store_true", help="关闭实体归一化（消融实验用）")
    ap.add_argument("--global_mode", default="concat", choices=["concat", "roundrobin"],
                    help="GraphRAG 全局查询的社区文档组织方式：拼接截断 / 轮转交错")
    ap.add_argument("--systems", default="vector,graph,wiki",
                    help="参与评测的范式：vector,graph,wiki,hybrid 的逗号组合")
    ap.add_argument("--only", default="builtin+self",
                    choices=["builtin+self", "builtin", "self"], help="评测用哪部分问题")
    ap.add_argument("--loose", action="store_true", help="faith 模式用无「禁止编造」约束的对照 prompt")
    ap.add_argument("--out", default=None, help="结果 JSON 输出路径")
    args = ap.parse_args()

    if args.mrr_check:  # 手算复核：MRR = (1/3+1/1+1/2)/3 ≈ 0.611；排名 0 表示未命中，计 0 分
        ranks = [int(x) for x in args.mrr_check.split(",")]
        terms = [f"1/{r}" if r > 0 else "0（未命中）" for r in ranks]
        m = sum(1.0 / r if r > 0 else 0.0 for r in ranks) / len(ranks)
        print("MRR = (" + " + ".join(terms) + f") / {len(ranks)} = {m:.3f}")
        return

    if args.mode == "extract":
        extract_triples(args.out)
        return
    if args.mode == "entries":
        print_entries()
        return
    if args.mode == "triples":
        print_triples()
        return

    if args.only == "builtin":
        questions = [q for q in QUESTIONS if q["src"] == "内置"]
    elif args.only == "self":
        questions = [q for q in QUESTIONS if q["src"] == "自拟"]
    else:
        questions = QUESTIONS

    emb = Embedder(args.emb)
    all_systems = [x.strip() for x in args.systems.split(",") if x.strip()]

    if args.mode == "update":
        rep = update_experiment(emb)
        if args.out:
            dump_json(args.out, rep)
        return

    if args.mode == "scan":     # [修改6] K 敏感性扫描
        scan = {}
        print("\n===== K 敏感性扫描（内置 15 题，逐档重跑评测）=====")
        header = f"{'范式':<12}{'类型':<8}" + "".join(f"{'K=' + str(k):>9}" for k in (1, 2, 3, 5))
        print(header)
        for k in (1, 2, 3, 5):
            systems, _ = build_systems(emb, all_systems, args.doc16, hops=args.hops,
                                       comms=args.comms, normalize=not args.no_norm,
                                       global_mode=args.global_mode)
            rep = evaluate(systems, k, questions, quiet=True, scope=args.scope)
            scan[k] = rep
        seen_names = list(dict.fromkeys(r["paradigm"] for r in scan[1]["rows"]))
        for name in seen_names:
            for tp in ("fact", "multi", "global"):
                line = f"{name:<14}{TYPE_CN[tp]:<10}"
                for k in (1, 2, 3, 5):
                    row = next(r for r in scan[k]["rows"] if r["paradigm"] == name and r["type"] == tp)
                    line += f"{row['recall']:>9.3f}"
                print(line)
        if args.out:
            dump_json(args.out, {str(k): v for k, v in scan.items()})
        return

    systems, corpus = build_systems(emb, all_systems, args.doc16, hops=args.hops,
                                    comms=args.comms, normalize=not args.no_norm,
                                    global_mode=args.global_mode)

    if args.mode == "eval":
        rep = evaluate(systems, args.k, questions, scope=args.scope)
        if args.out:
            dump_json(args.out, rep)
        return

    if args.mode == "detail":
        rows = detail_report(systems, args.k, questions, args.scope)
        if args.out:
            dump_json(args.out, {"k": args.k, "rows": rows})
        else:
            print(json.dumps(rows, ensure_ascii=False, indent=2))
        return

    if args.mode == "faith":
        qids = [x.strip().upper() for x in args.q.split(",")]
        faith_report(systems, qids, args.out or os.path.join(ROOT, "out", "faith.json"),
                     loose=args.loose)
        return

    # 单题模式
    scope = args.scope if args.mode == "graph" else "local"
    if args.mode != "graph":                    # 只保留被请求的范式，明细更干净
        systems = {k: v for k, v in systems.items() if k.lower().startswith(args.mode.split("_")[0])} \
            if args.mode in ("vector", "wiki", "hybrid") else systems
    q = next(q for q in QUESTIONS if q["id"].upper() == args.q.upper())
    print_single(systems, q, args.k, scope, emb.ok)


if __name__ == "__main__":
    main()
