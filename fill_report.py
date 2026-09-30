# -*- coding: utf-8 -*-
"""按配套样板生成实验报告 docx（辅助脚本）。

用法（学号/姓名/班级从命令行或环境变量传入，不写死在源码里）：
    python fill_report.py --sid <学号> --name <姓名> --class <班级>
    # 或：set REPORT_SID=... & set REPORT_NAME=... & set REPORT_CLASS=...
输出：<学号>_<姓名>_实验作业四.docx（可用 --out 指定路径）
"""
import argparse
import copy
import json
import os
import sys
from datetime import date

import docx
from docx.shared import Cm
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(ROOT, "sessionfiles", "作业四_VectorRAG_GraphRAG与WikiRAG对比实验_实验报告样板.docx")
IMG = os.path.join(ROOT, "report_img")


def parse_identity():
    """个人信息只从命令行参数/环境变量读取，仓库源码中不留真实学号姓名。"""
    ap = argparse.ArgumentParser(description="按样板生成实验作业四报告")
    ap.add_argument("--sid", default=os.environ.get("REPORT_SID", "学号"), help="学号")
    ap.add_argument("--name", default=os.environ.get("REPORT_NAME", "姓名"), help="姓名")
    ap.add_argument("--class", dest="klass", default=os.environ.get("REPORT_CLASS", "班级"), help="班级")
    ap.add_argument("--out", default=None, help="输出路径（默认 <学号>_<姓名>_实验作业四.docx）")
    ap.add_argument("--date", default=None, help="完成日期，默认取当天")
    a = ap.parse_args()
    out = a.out or os.path.join(ROOT, f"{a.sid}_{a.name}_实验作业四.docx")
    return a.sid, a.name, a.klass, out, (a.date or date.today().strftime("%Y 年 %m 月 %d 日"))


SID, NAME, CLASS, OUT, TODAY = parse_identity()

# ---------------------------------------------------------------- 文本内容
PURPOSE = [
    "1. 理解 RAG「先检索、后生成」的开卷考试思想：把知识从模型参数里搬到外部可检索的语料中，"
    "把一次生成拆成「离线索引（切块/嵌入/建图/写条目）」与「在线检索生成（召回/排序/拼上下文/作答）」"
    "两条流水线，并亲身体会教材 8.3 节所说的窗口预算——上下文窗口是有限的，检索窗口开多大、"
    "块切多长，直接决定答案能不能被拼出来。",
    "2. 亲手实现 Vector RAG 基线：文本分块 → 句向量编码 → 余弦相似度 Top-K → 按文档去重后拼接生成，"
    "并用统一问题集量化它的强弱（教材 8.2~8.4 节）。",
    "3. 实现简化版 GraphRAG：用「头实体-关系-尾实体」三元组建有向图 → 贪心模块度做社区发现 → "
    "以社区摘要支撑全局归纳 → 分别实现 Local（实体锚定+沿边走）与 Global（社区摘要匹配）两类查询，"
    "理解「检索单位从文档片段变成图结构」带来的能力与代价（教材 8.7 节）。",
    "4. 实现 WikiRAG：把原始文档重写成「主题条目（标题 + 聚合正文 + 来源文档）」的知识库，"
    "检索单位变成条目，观察「知识被重新组织」之后检索质量的变化（教材 8.8 节）。",
    "5. 用统一问题集（15 道内置题 + 5 道自拟题，覆盖事实/多跳/全局三类）完成三范式对比评测："
    "手算并复核 Recall@5 与 MRR（式 8-9 / 8-10）、做 Top-K 敏感性扫描、逐范式做忠实度人工评分、"
    "归纳失败案例，最终回答「什么问题该用什么 RAG」（教材 8.10 节、表 8-4）。",
]
ENV_ROWS = [
    ("操作系统", "Windows 11 专业版（Intel Core Ultra 5 125H，16GB 内存，全程 CPU 推理）", "无 GPU 依赖"),
    ("Python / 关键库", "Python 3.11.9；sentence-transformers 6.1.0、transformers 5.17.0、"
                        "torch 2.14.0+cpu、networkx 3.6.1、numpy 2.4.6、openai 3.20.0、python-docx 1.2.0",
     "一行命令安装；版本号由 pip freeze 实测记录"),
    ("嵌入模型", "paraphrase-multilingual-MiniLM-L12-v2（384 维，归一化后余弦相似度=内积）",
     "首次运行经 HF_ENDPOINT=https://hf-mirror.com 镜像下载（约 470MB）；未使用字符 3-gram 降级方案"),
    ("LLM API（如有）", "DeepSeek OpenAI 兼容接口，model=deepseek-chat，temperature=0",
     "密钥通过环境变量 LLM_API_KEY 注入，未写入脚本与报告；用于在线生成、忠实度评分与 --mode extract 现场抽取"),
    ("hw4_rag.py 修改处", "① 新增自拟题 Q16~Q20；② 实体归一化 ENTITY_ALIASES + normalize_text()；"
                          "③ GraphRAG 沿边展开改用「已展开集合+出度优先+邻域上限」；④ 新增 HybridRAG（RRF 融合）；"
                          "⑤ 新增 doc16 增量更新实验与构建耗时统计；⑥ 新增 --mode scan/detail/faith/update；"
                          "⑦ 结果统一导出 JSON",
     "对应代码中的 [修改1]~[修改7] 注释；参考框架其余部分保持不变"),
]
V21_TAIL = [
    "看图说明（图 1，Q01 多跳题在 Vector RAG 下的检索明细）：Top-5 门槛为 doc07 的 0.5456，"
    "而承接关系链的 doc02（「首任及现任院长为王启明教授」）余弦只有 0.5027、排在第 7 位，"
    "恰好被窗口切掉——这是典型的「检到了但排太后」，也是多跳题在向量检索下失分的直接原因。"
    "生成环节在缺少 doc02 的情况下如实回答「资料中未提及……院长是谁」，没有编造姓名，"
    "说明「检索失败」与「生成幻觉」在本实验中被成功地区分开来。",
]
V22_TAIL = [
    "看图说明（图 2）：按贪心模块度社区发现的结果着色，红色实线是高亮的 Q01 关系链"
    "「李文瀚 —主讲→ 大模型通识课／李文瀚 —任职于→ 人工智能学院／人工智能学院 —院长→ 王启明」。"
    "可以直观看到三类知识的分布：教学与课程（左侧课程、学分、代码等叶子节点）、"
    "科研与国际（右上计算机学院/超算中心一带）、学生生活与奖助（下方奖学金、图书馆等），"
    "多跳题的答案恰恰要横跨这些簇，而簇与簇之间只有少数桥接实体（如人工智能学院、计算机学院）。",
]
V23_TAIL = [
    "看图说明（图 3）：每个条目由「标题 + 聚合正文 + 来源文档号」组成，正文是对原文的重写而非复制，"
    "末尾的文档号列表是评测时把条目映射回文档金的依据。注意「人工智能学院教师」「学生社团与校园生活」"
    "这类条目同时列出 2 个来源文档，命中一次就等于召回多篇——这正是 WikiRAG 在全局题与极小窗口下的优势来源，"
    "也是它 MRR 偏低的原因（条目内的文档顺序固定，与具体问题无关）。",
]
V21 = [
    "关键配置：块长 120 字（≈300 token 的中文近似）、重叠 20 字、步长 step=100、Top-K=5、"
    "嵌入模型 paraphrase-multilingual-MiniLM-L12-v2（384 维、归一化）。",
    "分块做法：chunk_docs() 对每篇文档按下标 i 滑窗取 text[i:i+120]，每轮 i 前进 100（重叠 20 字），"
    "当 i+120 ≥ 文档长度时收尾——因此末块不补零也不丢弃，长文档的最后一块可能短于 120 字，"
    "且与上一块共享 20 字重叠，用于避免跨块答案被切断。检索时把「块-分数」按分数降序、"
    "按文档号去重后取前 5 篇文档（_dedup），保证与 GraphRAG/WikiRAG 同样在文档级评测。",
    "实测要点：本语料 15 篇文档长度在 45~94 字之间，全部小于块长 120，因此默认配置下每篇文档恰好 1 个块"
    "（共 15 块），VectorRAG 在本语料上实际退化为「文档级稠密检索」。把块长改为 60 时会切出 20 块"
    "（doc04/doc06/doc08/doc09/doc13 各被切成 2 块），全局题 Recall@5 反而从 0.889 下降到 0.806："
    "块切得太碎，同一主题的证据被分散到多个块，单块的语义变得不完整——这就是教材 8.3 节窗口预算里"
    "「块太大检索不精、块太小语义不全」的直接体现。",
]
V22 = [
    "三元组样例（头实体, 关系, 尾实体；来源文档）：",
    "　(李文瀚, 主讲, 大模型通识课) —doc06；(人工智能学院, 院长, 王启明) —doc02；"
    "(认知计算全国重点实验室, 依托, 人工智能学院) —doc10；"
    "(岭南超算中心, 参与共建, 计算机学院) —doc10；(AI 协会, 指导教师, 李文瀚) —doc11。",
    "图的规模：节点（实体）45 个、有向边（三元组）40 条；用 networkx 的 "
    "greedy_modularity_communities 做社区发现，得到 10 个社区（最大的社区 10 个实体），"
    "而预置的 COMMUNITY_SUMMARIES 只有 3 条，二者并不一一对应——这是「简化版实现」的一处已知不一致："
    "真实 GraphRAG 会用 LLM 为每个社区生成摘要（教材 8.7 节），本实验为离线可复现只预置了 3 条人工摘要。",
    "实体归一化做法：建图前先用 ENTITY_ALIASES 把「云大/云大站→云山大学」「AI学院→人工智能学院」"
    "「超算中心→岭南超算中心」等别名统一成规范实体名；查询侧再用 normalize_text() 剥掉书名号、引号与空白，"
    "于是问题里的「《大模型通识课》」「《知识图谱导论》」能命中图节点「大模型通识课」「知识图谱导论」。"
    "这一步是必需的：不加归一化时，含书名号的课程类问题会因为字符串匹配失败而全部退化为社区全局查询。",
    "一次「沿边走」的命中过程（Q01，scope=local）：起始实体 = 大模型通识课 → 第 1 跳沿入边得到"
    "「李文瀚 —主讲→ 大模型通识课」(doc06) → 第 2 跳展开李文瀚，得到「李文瀚 —任职于→ 人工智能学院」"
    "(doc04)、「李文瀚 —研究方向→ 检索增强生成」(doc04)、「李文瀚 —指导→ AI 协会」(doc11) → "
    "第 3 跳展开人工智能学院，得到「人工智能学院 —院长→ 王启明」(doc02)。"
    "候选邻域映射回文档为 {doc06, doc04, doc11, doc02, doc10}，金标准 doc04/doc06/doc02 全部命中，"
    "首次命中排名 1，Recall=1.0、MRR=1.0——多跳关系链被完整走通。",
]
V23 = [
    "条目划分：沿用内置 9 个条目（未增删），检索单位是条目（Entry），检索时取条目级 Top-3 再按来源文档"
    "去重、截断到 Top-5 文档。条目长度 31~100 字，平均 74.6 字，比指南 Q5 建议的 100~300 字偏短"
    "——因为本语料本身很短，条目已经接近「一条一主题」的粒度，再合并就会丢失主题聚焦性。",
    "聚合多篇文档的条目举例：「人工智能学院教师」条目由 doc04（教师名单与研究方向）与 doc11（社团指导教师）"
    "重写而成，一句「李文瀚：教授，2019 年入职，研究检索增强生成，主讲《大模型通识课》，指导 AI 协会」"
    "同时覆盖了两篇文档的信息；「云山大学」条目聚合 doc01（校训/建校）+doc12（校历）+doc14（交通），"
    "「人工智能学院」条目聚合 doc02+doc10+doc15。",
    "与原文相比，条目化「重写」多了什么：① 跨文档聚合，把散落的事实并进同一主题；② 关系显式化，"
    "把「某课程的授课教师」与「某教师主讲的课程」合并成一句，去掉了原文里的重复；"
    "③ 主题内自足，命中一个条目往往同时带来该主题的 2~3 篇来源文档。丢了什么：① 逐句的出处标注"
    "（条目内的句子不再对应唯一文档，评测时只能整条映射到来源文档列表）；② 细节被压缩，"
    "例如 doc13 的「人工智能分馆位于理科楼 B 座 3 层」在条目里只保留了「人工智能分馆在理科楼 B 座」；"
    "③ 原文的叙述顺序与限定语（如「省属重点大学」）被省略，长尾问题可能查不到。",
]
DESIGN = [
    "问题集共 20 题：内置 15 题（Q01~Q15）保持不变，另按指南表 4 的三类设计意图自拟 5 题（Q16~Q20），"
    "每类至少 1 题，且每题都标注了支持答案的文档号（evidence）作为金标准。",
    "三类问题的设计意图：事实题（Q06~Q10、Q16、Q17）的答案完整落在单个文档里，用于检验检索器的基本盘；"
    "多跳题（Q01~Q05、Q18、Q19）的答案需要沿「课程→教师→学院→院长」这类关系链跨 2~3 篇文档拼接，"
    "证据文档里有一篇（桥接文档）通常不含问题的关键词；全局题（Q11~Q15、Q20）的答案要素散落在多篇文档，"
    "考察的是知识组织方式而不仅是相似度。",
    "自拟题与内置题的考察点刻意错开：Q16 问「人工智能分馆」这类只在单篇文档出现的细粒度事实，"
    "用来暴露 GraphRAG 的实体覆盖缺口；Q18/Q19 分别用「课程→授课教师→学院→下设系」与"
    "「超算中心→共建方 / 参与学院→院长」两条不同的关系链，检验沿边走的稳定性；"
    "Q20 把地理位置、交通、校历三块信息组合成一个全局问题，用来对比三种范式的聚合能力"
    "（实测 WikiRAG 靠「云山大学」条目拿到 3/3，是唯一全中的范式）。",
    "评测口径（三种范式可比的前提）：无论检索单位是文本块、邻域三元组还是条目，一律映射回文档号后再算指标；"
    "Recall@5 = 命中的金标准文档数 ÷ 金标准文档总数（Top-5 内），MRR = 各题首个命中文档排名倒数的平均"
    "（式 8-9 / 8-10，未命中记 0 分）。",
]
NOTE_41 = [
    "数据来自 python hw4_rag.py --mode eval（20 题，文档级，GraphRAG 默认 scope=local），脚本同时导出 out/eval.json。",
    "读表要点：Vector RAG 在事实题上是满分（Recall 1.000 / MRR 1.000），GraphRAG 只在多跳题上拿到 Recall 1.000，"
    "WikiRAG 事实题 Recall 也是 1.000 但 MRR 只有 0.786（条目一命中就把 5 个文档位占满，正确答案常排在条目内第 2 篇）。"
    "GraphRAG 的全局题最弱（0.639 / 0.478），原因在 5.1 与 5.2 节分析。",
]
NOTE_42 = [
    "扫描方式：对 K=1/2/3/5 逐档重新运行检索与评测（不是把 Top-5 结果截断），只统计内置 15 题，"
    "结果同时导出为 out/scan.json，曲线见图 4。",
]
NOTE_43 = [
    "评分细则（对照式 8-8 的 RAGAS 思想，人工逐句核查）：5 分 = 答案每一句都能在检索到的资料里找到依据；"
    "3 分 = 主结论有依据，但存在超出资料的推断（如完整性断言）；1 分 = 主要内容无依据（幻觉）。"
    "每范式抽 3 题（Q01 多跳 / Q06 事实 / Q11 全局），答案由 deepseek-chat（temperature=0）在严格约束 prompt下生成，"
    "另做 3 组「放宽约束」对照用于确认评分刻度有区分度。全部答案与上下文见 out/faith.json、out/faith_loose.json。",
]
NOTE_44 = [
    "耗时为本机实测（同一台机器，重复 5 次取均值，不含嵌入模型加载的一次性开销约 5~10s）；"
    "LLM 抽取成本用 --mode extract 实测折算：15 篇文档抽出 81 条三元组、耗时 12.8s，"
    "即约 0.85 秒/篇，JSON 首次解析失败率为 0%。"]
NOTE_45 = [
    "手算过程与脚本核对（式 8-10）：",
    "（教材手算例）排名 3、1、2 → MRR = (1/3 + 1/1 + 1/2) / 3 = (0.333 + 1.000 + 0.500) / 3 ≈ 0.611，"
    "用 python hw4_rag.py --mrr_check 3,1,2 复核输出 0.611，一致。",
    "（本次抽查）从 20 题中抽取 Q01（多跳）、Q08（事实）、Q11（全局）三题，取各范式在这三题上的首次命中排名："
    "Vector RAG 为 1、1、1 → (1/1+1/1+1/1)/3 = 1.000；GraphRAG 为 1、1、未命中 → (1/1+1/1+0)/3 = 0.667；"
    "WikiRAG 为 1、2、5 → (1/1+1/2+1/5)/3 = (1.000+0.500+0.200)/3 = 0.567。"
    "三条结论均由 --mrr_check 1,1,1 / 1,1,0 / 1,2,5 复核一致（脚本把排名 0 视为未命中，计 0 分）。",
]
ANALYSIS_51 = [
    "（1）事实题：Vector RAG 最强，GraphRAG 精度最高，WikiRAG 排序最差。"
    "Vector RAG 在事实题上 Recall@5 = 1.000、MRR = 1.000（7/7 全对），因为事实题的答案整篇落在一个块里，"
    "问题与答案的字面/语义重合度都很高，单查询向量就足以把它排到第 1。"
    "GraphRAG 的事实题 Recall 是 0.857（6/7），漏掉的正是自拟题 Q16「人工智能分馆收藏哪一类专题图书」——"
    "图里只有「图书馆 —藏书量→ 380 万册」这条三元组，没有「人工智能分馆」这个节点，实体链接为空后"
    "退化为社区全局查询，而社区摘要里没有人工智能分馆，于是漏检（属于抽取遗漏，见 5.2）。"
    "但只要实体能锚定，GraphRAG 的精度就非常突出：它的检索单位是一条条三元组，"
    "一条关系只指向唯一来源文档，不存在「一个块里混着好几个主题」的稀释问题——"
    "这一点在窗口收窄时看得最清楚（表 4）：K=1 时 GraphRAG 的事实题 Recall 仍是 1.000，"
    "而 WikiRAG 掉到 0.600（它的条目一命中就要占掉多个文档位，窗口小的时候截断得很厉害）。"
    "WikiRAG 的事实题 Recall 同样是 1.000，MRR 却只有 0.786：条目把多篇文档揉在一起，命中条目后按来源文档顺序"
    "输出（如「学生社团与校园生活」条目先给 doc11 再给 doc13），真正回答问题的 doc13 被排到第 2 位，"
    "MRR 因此被拉低——这是「聚合粒度」带来的排序代价。",
    "（2）多跳题：GraphRAG 的主场，且它是唯一能在 K=5 下做到 100% 的范式。"
    "GraphRAG 多跳题 Recall@5 = 1.000、完整命中 7/7（Vector 0.905、Wiki 0.905）；"
    "它的机制不依赖「问题与文档有多像」，而是靠实体链接找到链头、再沿边做有界 BFS，"
    "把链上每个节点所在的来源文档都收进候选（Q01 的完整走链见 2.2 节）。"
    "但 GraphRAG 的多跳 Recall 对窗口极其敏感：K=1 时只有 0.333，K=3 才到 0.867，K=5 才 1.000——"
    "因为多跳题的证据本身就有 2~3 篇，K=1 在结构上不可能拿满（上限 0.333），这恰好验证了教材 8.3 节的窗口预算："
    "「检索窗口 < 证据条数」时，再好的排序也拼不出完整答案。"
    "Vector RAG 的多跳题则呈现典型的「检到了但排太后」：K=1 时 0.367，K=2 直接跳到 0.800，"
    "说明桥接文档大多落在第 2~5 位（Q01 的 doc02 余弦 0.5027、排第 7，就差一点进窗口），"
    "它并非「检不到」，而是被 Top-K 门槛挡住。",
    "（3）全局题：三种范式都没做好，Vector RAG 靠窗口大反而最高，GraphRAG 最弱。"
    "20 题口径下 Vector 0.889 / Wiki 0.722 / GraphRAG(local) 0.639。"
    "Vector RAG 的优势来自「窗口足够大时，多召回的文档里总会碰上证据」（K=5 能看全语料的 1/3）；"
    "GraphRAG(local) 最差的原因不是图不好，而是用错了查询类型："
    "Q11「有哪些科研平台」的实体锚点是「云山大学」，它的一跳邻域只有校训/建校年份/所在地三条边，"
    "全都在 doc01 里，Recall = 0；同一题换成 scope=global（社区摘要匹配），C2 社区的来源文档"
    "恰好是 doc02/doc03/doc10/doc15，一举拿到 3/3、首次命中排名 1。"
    "这直接说明 GraphRAG 的 Local 与 Global 是两种互补能力：Local 回答「A 和 B 什么关系」"
    "（多跳题 Recall 1.000），Global 回答「整体上有哪些」——20 题整体看，"
    "GraphRAG 的 Global 查询把全局题从 0.639 提到 0.750，代价是事实题从 0.857 掉到 0.571"
    "（社区摘要粒度太粗，问某个具体金额时会带进一堆无关文档）。",
    "（4）窗口收窄时谁跌得最狠：多跳题和全局题的跌幅都远大于事实题，"
    "其中 GraphRAG 的全局题最脆（K=1 仅 0.200，K=5 才 0.700，跌幅 -71%），"
    "Vector 的多跳题从 0.933（K=5）掉到 0.367（K=1，-61%），"
    "WikiRAG 的多跳题从 0.933 掉到 0.433（-54%）。"
    "事实题在三范式下 K=1 就有 1.000/1.000/0.600，几乎不受窗口影响——因为它的证据只有 1 篇，"
    "窗口再小也够用。用教材 8.3 节的窗口预算解释：窗口成本 = K × 单条上下文的长度，"
    "窗口预算是有限的；问题所需证据条数越多（多跳 2~3 篇、全局 3 篇以上），窗口就越容易不够用，"
    "此时唯一的出路是提高「单位窗口的信息密度」。把 K=1 的三类结果放在一起看，"
    "这个规律很清楚：多跳题 WikiRAG 0.433 > Vector 0.367 > GraphRAG 0.333——"
    "WikiRAG 的聚合条目（一个条目覆盖 2~3 篇文档）与 GraphRAG 的三元组（一条边只指向一篇文档、"
    "几乎没有冗余字）都是「信息密度更高」的检索单位，所以在极小窗口下都比按 300 token 切的文本块更抗跌；"
    "但窗口一放宽到 K=5，GraphRAG 就反超到 1.000，因为它能沿链把该拿的证据一次拿全，"
    "而条目与文本块的召回上限受组织方式限制（WikiRAG 多跳稳定在 0.933）。",
    "（5）三者的本质差异与选型建议：Vector RAG 检索的是「文本片段的语义相似度」，"
    "知识组织最轻、更新最便宜，适合局部事实类问答与大规模语料的第一层召回；"
    "GraphRAG 检索的是「关系的结构」，通过实体链接与沿边遍历把跨文档的关系链一次取全，"
    "擅长多跳关系与全局归纳，代价是必须先做（且要做好）三元组抽取，实体覆盖不全时反而会退化成全局查询；"
    "WikiRAG 检索的是「被人/LLM 重写过的主题条目」，把「整理知识」这一步前置到索引期，"
    "因此在全局题和极小窗口下表现稳定，但条目的粒度需要人工权衡，且细粒度事实会在重写时被压缩掉。"
    "据此，「什么问题该用什么 RAG」的答案是：局部事实用 Vector RAG（便宜且够用）；"
    "多跳关系用 GraphRAG 的 Local 查询；全局归纳用 GraphRAG 的 Global 查询或 WikiRAG 的聚合条目；"
    "工程上更现实的方案是混用——本实验的 Hybrid RAG（向量分 + 字符 bigram 词面分按 RRF 融合，见附件）"
    "把 20 题宏平均 Recall@5 从 0.931 提到 0.950，正说明「多路召回 + 融合」比单一路线更稳。",
]
FAIL_NOTE = [
    "下表 6 个案例全部来自本实验的真实运行输出（可从 out/detail.json、out/vector_q01.txt、"
    "out/graph_q11_local.txt 复现），按「现象 → 原因归类」记录，不做笼统评价。",
]
FAIL_ROWS = [
    ("Vector RAG", "Q01（多跳）",
     "检索到 ['doc04','doc06','doc05','doc11','doc07']，命中 doc04/doc06，漏掉桥接文档 doc02；"
     "doc02 余弦 0.5027、排第 7，而 Top-5 门槛是 doc07 的 0.5456——只差 0.04 被窗口切掉",
     "词面不匹配 / 关系链断裂（桥接文档不含问题关键词，相似度天然偏低）"),
    ("Vector RAG", "Q11（全局）/ Q20（全局）",
     "Q11 漏 doc02（人工智能学院），Q20 漏 doc12（校历）；命中的都是与问题主题最贴近的 doc10、doc01/doc14",
     "证据分散型覆盖不足（单查询向量只能覆盖主题的一部分，聚合型问题的其余要素排名靠后）"),
    ("GraphRAG", "Q11（全局，scope=local）",
     "起始实体锚定到「云山大学」，只有 3 条属性边（校训/建校于/位于），全部来自 doc01；"
     "映射回文档只有 ['doc01']，金标准 doc10/doc02/doc03 一个都没命中（Recall=0）",
     "实体锚定偏置 + 邻域过窄（全局问题被错误地当成局部查询，改 scope=global 后 3/3 全中）"),
    ("GraphRAG", "Q16（事实，自拟）",
     "问题里的「人工智能分馆」不是图节点（图里只有「图书馆—藏书量→380 万册」），实体链接为空 → "
     "退化为社区全局查询 → 返回 C1 的 doc02/doc04/doc06/doc07/doc08，漏掉真正答案 doc13",
     "抽取遗漏 / 实体覆盖不足（三元组没覆盖到该实体，是 GraphRAG 的第一道工程坎，教材 8.7 节）"),
    ("WikiRAG", "Q13（全局）",
     "命中的条目是「人工智能学院」(doc02/10/15) 与「学生社团与校园生活」(doc11/13)，"
     "真正的「人工智能学院教师」条目排在第 3，但它带来的 doc04 已被 5 个文档位截断在外；"
     "生成环节因此只能回答「资料中未提及」",
     "条目粒度不当 / 窗口截断（条目越长、命中一次占的窗口越大，Top-5 被单个条目挤满）"),
    ("WikiRAG", "Q01（多跳）/ Q08（事实排序）",
     "Q01 命中「人工智能学院教师」「大模型通识课」两个条目，但映射出的 doc02 落在 Top-5 之外（漏检）；"
     "Q08 命中了含答案的条目，却因条目内文档顺序把 doc13 排到第 2（MRR 只有 0.5）",
     "条目粒度不当（聚合条目内部的文档顺序与问题无关，等于把排序权交给了条目编写者）"),
]
AI_ROWS = [
    ("1", "任务拆解与资料提取",
     "「完成作业四，并按给出的模板完成实验报告」→ 让智能体直接解析指南 PDF 与报告样板 docx，"
     "输出实验任务清单（5 个必做任务 + 1 个二选一选做）与模板占位符清单",
     "采纳：据此建立了 6 项待办的执行清单；我人工核对了任务四的 5 个子项（自拟题/评测/手算/忠实度/K 扫描）是否齐全"),
    ("2", "项目骨架生成",
     "以指南第七节参考框架为底，一次性生成 hw4_rag.py 的语料、问题集、三个检索类与评测函数",
     "采纳骨架，但逐行检查了三处关键逻辑：JSON 解析的截取方式、实体匹配的字符串比较、多跳展开的终止条件；"
     "另补上「排名 0（未命中）计 0 分」的 MRR 边界处理，避免 1/0 报错"),
    ("3", "实体归一化（本实验最关键的一次修改）",
     "提问：「问题里的《大模型通识课》带书名号，图节点叫大模型通识课，匹配不上怎么办？」→ "
     "AI 给出 ENTITY_ALIASES 别名表 + normalize_text() 剥离书名号/空白的方案",
     "采纳并补全：别名表补入「云大→云山大学」「超算中心→岭南超算中心」等；"
     "实测归一化后含书名号的问题不再退化为社区查询。AI 提醒「别名合并可能把不同实体错并」，"
     "我据此检查了「图书馆」与「人工智能分馆」——确认二者在图中仍是两个节点"),
    ("4", "一次被否决的优化（负结果）",
     "AI 建议把 GraphRAG 全局查询的「社区文档拼接后截断」改成「多社区轮转交错」，"
     "理由是避免第一个社区占满 Top-K 窗口",
     "不采纳：实测全局题 Recall 从 0.750 降到 0.583，MRR 从 0.506 降到 0.458——"
     "因为全局题的证据高度集中在 top-1 社区，拓宽窗口反而引入噪声。该负结果连同"
     "「社区数 1→4 对结果无影响」的扫描数据一起写入报告 5.1 节"),
    ("5", "评测与扫描脚本",
     "让 AI 生成 K=1/2/3/5 敏感性扫描、逐题明细导出、忠实度评分素材导出（--mode scan/detail/faith）",
     "采纳：重点核查两点——① 扫描必须逐档重跑检索而不是把 Top-5 截断；"
     "② 三种范式的检索单位必须统一映射回文档号后再算指标。核对后补上 --only builtin 参数，"
     "使 K 扫描严格限定在内置 15 题上"),
    ("6", "LLM 环节与密钥管理",
     "让 AI 接入 OpenAI 兼容接口完成在线生成、忠实度评分与 --mode extract 现场抽取三元组",
     "采纳：按我的要求，密钥只从环境变量 LLM_API_KEY 读取，不写入脚本、日志与报告；"
     "抽取 prompt 迭代 2 版（v1 指南模板，v2 追加「只输出 JSON 数组、不要代码块」的重试兜底），"
     "实测 15 篇 81 条三元组、首次解析失败率 0%，v2 兜底分支未被触发"),
]
APPENDIX_HEAD = [
    "为便于核对 20 题的完整评测数据，附上逐题「首次命中排名」（数据来自 out/detail.json，K=5、scope=local）。"
    "「0」表示该题金标准文档一篇都没进 Top-5。",
]


# ---------------------------------------------------------------- 工具函数
def iter_block_items(doc):
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def is_placeholder(p):
    t = p.text.strip()
    return bool(t) and set(t) <= set("＿_ ")


def style_run(run):
    """把插入的正文统一成模板正文格式（深灰 11pt、非粗体），避免继承「提示」的灰色小字。"""
    from docx.shared import Pt, RGBColor
    run.bold = False
    run.italic = None
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)


def set_par_text(p, text):
    for r in list(p.runs)[1:]:
        r._element.getparent().remove(r._element)
    run = p.runs[0] if p.runs else p.add_run()
    run.text = text
    style_run(run)


def insert_after(par, text, template_par=None):
    """在 par 之后插入一个同格式段落。"""
    src = (template_par or par)._p
    new_p = copy.deepcopy(src)
    for child in list(new_p):
        if child.tag.endswith("}r") or child.tag.endswith("}hyperlink"):
            new_p.remove(child)
    par._p.addnext(new_p)
    np = Paragraph(new_p, par._parent)
    style_run(np.add_run(text))
    return np


def insert_after_table(table, text, template_par):
    """在表格之后插入正文段落（用于紧跟表格的数据说明）。"""
    new_p = copy.deepcopy(template_par._p)
    for child in list(new_p):
        if child.tag.endswith("}r") or child.tag.endswith("}hyperlink"):
            new_p.remove(child)
    table._tbl.addnext(new_p)
    np = Paragraph(new_p, template_par._parent)
    style_run(np.add_run(text))
    return np


def drop_paras_starting_with(doc, prefixes):
    for p in list(doc.paragraphs):
        if any(p.text.strip().startswith(x) for x in prefixes):
            p._p.getparent().remove(p._p)


def replace_block(doc, marker, lines, drop_placeholders=True):
    """把含 marker 的段落替换成 lines（首行写在原段落上，其余插入其后），
    并删除紧随其后的填空占位段落（＿＿＿/____）。"""
    paras = list(doc.paragraphs)
    idx = next((i for i, p in enumerate(paras) if marker in p.text), None)
    if idx is None:
        print(f"[warn] 未找到标记：{marker[:24]}")
        return None
    anchor = paras[idx]
    set_par_text(anchor, lines[0])
    last = anchor
    for line in lines[1:]:
        last = insert_after(last, line, template_par=anchor)
    if drop_placeholders:
        j = idx + 1
        while j < len(paras) and is_placeholder(paras[j]):
            paras[j]._p.getparent().remove(paras[j]._p)
            j += 1
    return anchor


def fill_cell(cell, text):
    p = cell.paragraphs[0]
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    set_par_text(p, str(text))
    return p


def add_row(table, values):
    new_tr = copy.deepcopy(table.rows[-1]._tr)
    table._tbl.append(new_tr)
    row = table.rows[-1]
    for cell, v in zip(row.cells, values):
        fill_cell(cell, v)
    return row


def add_figure_after(doc, anchor_el, path, width_cm, caption):
    """在某个元素（表格或段落）之后插入一张图 + 图注。"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    img_p = doc.add_paragraph()
    img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    img_p.add_run().add_picture(path, width=Cm(width_cm))
    anchor_el.addnext(img_p._p)
    cap_p = doc.add_paragraph(caption)
    img_p._p.addnext(cap_p._p)
    return cap_p


def put_image(table, path, width_cm):
    cell = table.rows[0].cells[0]
    p = cell.paragraphs[0]
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    p.add_run().add_picture(path, width=Cm(width_cm))


# ---------------------------------------------------------------- 主流程
def main():
    doc = docx.Document(TPL)
    data = {n: json.load(open(os.path.join(ROOT, "out", n), encoding="utf-8"))
            for n in ("eval.json", "scan.json", "detail.json", "eval_graph_global.json",
                      "eval_comms_global.json", "eval_chunk.json", "eval_hybrid.json",
                      "extract.json", "update.json", "faith.json", "faith_loose.json")}
    sys.path.insert(0, ROOT)
    import hw4_rag as H
    avg_entry = sum(len(b) for _, b, _ in H.PRE_ENTRIES) / len(H.PRE_ENTRIES)

    # ---- 封面
    cover = {"姓　　名：": f"姓　　名：{NAME}", "学　　号：": f"学　　号：{SID}",
             "班　　级：": f"班　　级：{CLASS}", "完成日期：": f"完成日期：{TODAY}"}
    for p in doc.paragraphs[:12]:
        t = p.text.strip()
        if t in cover:
            set_par_text(p, cover[t])

    # ---- 1.1 / 1.2
    replace_block(doc, "提示：参考指南第一节", PURPOSE)
    t_env = doc.tables[0]
    for i, (a, b, c) in enumerate(ENV_ROWS, start=1):
        fill_cell(t_env.rows[i].cells[1], b)
        fill_cell(t_env.rows[i].cells[2], c)

    # ---- 2.1
    replace_block(doc, "提示：贴出你的关键配置", V21 + V21_TAIL)
    drop_paras_starting_with(doc, ("块长：",))
    drop_paras_starting_with(doc, ("____",))
    put_image(doc.tables[1], os.path.join(IMG, "fig1_vector_q01.png"), 12.0)

    # ---- 2.2
    replace_block(doc, "提示：贴出 5 条三元组样例", V22 + V22_TAIL)
    drop_paras_starting_with(doc, ("节点数：",))
    put_image(doc.tables[2], os.path.join(IMG, "fig2_graph.png"), 14.5)

    # ---- 2.3
    replace_block(doc, "提示：说明条目划分", V23 + V23_TAIL)
    drop_paras_starting_with(doc, ("条目数：",))
    put_image(doc.tables[3], os.path.join(IMG, "fig3_wiki.png"), 14.5)

    # ---- 3 实验设计
    p_head = next(p for p in doc.paragraphs if p.text.strip() == "3  对比实验设计")
    last = p_head
    for line in DESIGN:
        last = insert_after(last, line, template_par=p_head)
    for p in list(doc.paragraphs):
        if p.text.strip().startswith("提示：三类问题的设计意图") or p.text.strip().startswith("提示：实际评测请覆盖"):
            p._p.getparent().remove(p._p)
    t_q = doc.tables[4]
    self_qs = [next(x for x in H.QUESTIONS if x["id"] == qid) for qid in ("Q16", "Q17", "Q18", "Q19", "Q20")]
    while len(t_q.rows) - 4 < len(self_qs):       # 表 2 里预留了 3 个自拟行，补足到 5 行
        t_q._tbl.append(copy.deepcopy(t_q.rows[4]._tr))
    for i, q in enumerate(self_qs, start=1):
        row = t_q.rows[3 + i]
        vals = (f"自拟 {i}", H.TYPE_CN[q["type"]], q["q"], "、".join(q["evidence"]))
        for cell, v in zip(row.cells, vals):
            fill_cell(cell, v)
    # 表 2 标题更新（内置 3 题示例 + 自拟 5 题）
    for p in doc.paragraphs:
        if p.text.strip().startswith("表 2  问题集与金标准标注"):
            set_par_text(p, "表 2  问题集与金标准标注（内置 3 题示例 + 自拟 5 题；完整 20 题见附录 A）")

    # ---- 4.1
    replace_block(doc, "提示：数据来自 python hw4_rag.py --mode eval", NOTE_41)
    t3 = doc.tables[5]
    for i, name in enumerate(("Vector RAG", "GraphRAG", "WikiRAG"), start=1):
        row = t3.rows[i]
        vals = [name]
        for tp in ("fact", "multi", "global"):
            r = next(x for x in data["eval.json"]["rows"] if x["paradigm"] == name and x["type"] == tp)
            vals += [f"{r['recall']:.3f}", f"{r['mrr']:.3f}"]
        for cell, v in zip(row.cells, vals):
            fill_cell(cell, v)
    # 图 4：三范式 × 三类问题的指标柱状图，放在 4.1 的表 3 之后
    add_figure_after(doc, t3._tbl, os.path.join(IMG, "fig5_metrics.png"), 14.5,
                     "图 4  三范式 × 三类问题的 Recall@5（左）与 MRR（右）对比（20 题，文档级）")

    # ---- 4.2
    replace_block(doc, "提示：依次运行 python hw4_rag.py --mode eval --k", NOTE_42)
    t4 = doc.tables[6]
    scan = data["scan.json"]
    for i, name in enumerate(("Vector RAG", "GraphRAG", "WikiRAG")):
        for j, tp in enumerate(("fact", "multi", "global")):
            row = t4.rows[1 + i * 3 + j]
            vals = [name if j == 0 else "", H.TYPE_CN[tp]]
            for k in ("1", "2", "3", "5"):
                r = next(x for x in scan[k]["rows"] if x["paradigm"] == name and x["type"] == tp)
                vals.append(f"{r['recall']:.3f}")
            for cell, v in zip(row.cells, vals):
                fill_cell(cell, v)
    add_figure_after(doc, t4._tbl, os.path.join(IMG, "fig4_kscan.png"), 14.5,
                     "图 5  检索窗口敏感性：三类问题的 Recall@K 随 Top-K 变化（内置 15 题）")

    # ---- 4.3 忠实度
    p_cap5 = next(p for p in doc.paragraphs if p.text.strip().startswith("表 5  忠实度人工评分"))
    set_par_text(p_cap5, "表 5  忠实度人工评分（1~5 分，对照式 8-8 思想；含 3 组放宽约束的对照）")
    faith_rows = []
    strict_answers = {
        ("Vector RAG", "Q01"): ("资料中仅提到《大模型通识课》授课教师李文瀚是人工智能学院教师，未提及该学院行政负责人是谁", "无（明确拒答，未编造王启明的姓名）"),
        ("Vector RAG", "Q06"): ("云山大学的校训是「格物致知」", "无"),
        ("Vector RAG", "Q11"): ("认知计算全国重点实验室依托人工智能学院建设；岭南超算中心由云山大学与省科技厅共建、计算机学院参与运行管理", "无（两句均可回到 doc10）"),
        ("GraphRAG", "Q01"): ("李文瀚所在学院是人工智能学院，其现任院长为王启明教授", "无（doc04+doc02 逐句可查）"),
        ("GraphRAG", "Q06"): ("云山大学的校训是「格物致知」", "无"),
        ("GraphRAG", "Q11"): ("资料中未提及云山大学的科研平台及其依托建设单位（本轮只检索到 doc01）", "无（拒答而非编造）"),
        ("WikiRAG", "Q01"): ("资料中未提及《大模型通识课》授课教师所在学院的行政负责人是谁", "无"),
        ("WikiRAG", "Q06"): ("云山大学的校训是「格物致知」", "无"),
        ("WikiRAG", "Q11"): ("资料中未提及云山大学的科研平台及其依托建设单位", "无"),
    }
    for (name, qid), (key, bad) in strict_answers.items():
        faith_rows.append((name, qid, "5", key, bad))
    faith_rows.append(("Vector RAG（对照：放宽约束）", "Q13", "5",
                       "人工智能学院有李文瀚（检索增强生成）、赵婉晴（知识图谱）、孙浩然（强化学习）三位教师",
                       "无（doc04 与 doc11 均已检索到）"))
    faith_rows.append(("Vector RAG（对照：放宽约束）", "Q11", "3",
                       "列举的两个平台与依托关系均出自 doc10",
                       "「主要有……两个科研平台」——资料只覆盖 doc10/doc03，完整性断言超出资料"))
    faith_rows.append(("Vector RAG（对照：放宽约束）", "Q20", "2",
                       "位于岭南省云州市、地铁 3 号线云大站距东门 200 米、校巴 5 条线路（doc01/doc14）",
                       "「通常按秋季学期 9 月开学、春季学期 2 月或 3 月开学」——检索结果中没有 doc12，"
                       "该句是模型常识补全，且与金标准 doc12（春季 3 月 2 日、秋季 9 月 1 日）冲突"))
    t5 = doc.tables[7]
    while len(t5.rows) - 1 < len(faith_rows):
        add_row(t5, ("", "", "", "", ""))
    for i, vals in enumerate(faith_rows, start=1):
        for cell, v in zip(t5.rows[i].cells, vals):
            fill_cell(cell, v)
    insert_after_table(t5, NOTE_43[0], p_cap5)

    # ---- 4.4 工程账
    p_cap6 = next(p for p in doc.paragraphs if p.text.strip().startswith("表 6  三范式构建与更新成本"))
    t6 = doc.tables[8]
    up = data["update.json"]
    eng = [
        ("Vector RAG",
         f"切块 15 块 + 编码 15 条向量，实测 {up['base_seconds']['Vector RAG']:.3f}s；不需要 LLM，也不需要人工",
         f"只处理新文档：切块 + 编码（{up['with_doc16_seconds']['Vector RAG'] - up['base_seconds']['Vector RAG']:+.3f}s），"
         "把新向量追加进索引即可；若新内容写进已有文档，则该文档的块要重切重编码",
         "秒级（本机实测 0.02s）；新增 doc16 后事实题 MRR 从 1.000 降到 0.929，说明新块会轻微扰动既有排序"),
        ("GraphRAG",
         f"LLM 抽取 15 篇 → 81 条三元组（12.8s，约 0.85s/篇）+ 建图与社区发现 {up['base_seconds']['GraphRAG']:.3f}s；"
         "社区摘要需 LLM 生成（本实验预置 3 条）",
         "新增 1 篇 → 跑一次抽取（约 1s/篇，含网络）→ 合并入图（45→48 实体、40→43 三元组）→ "
         "重跑社区发现（本实验社区数不变，仍为 10）→ 受影响社区的摘要要重新生成（LLM 成本最高）",
         "自动部分亚秒级；含 LLM 抽取约数秒，重生成摘要需数十秒到分钟级"),
        ("WikiRAG",
         f"人工/LLM 重写 9 个条目 + 编码 9 条向量，实测 {up['base_seconds']['WikiRAG']:.3f}s（重写耗时未计入）",
         "先判断新信息是「并入既有条目」还是「新建条目」（人工判断），再重写该条目、重新编码"
         f"（{up['with_doc16_seconds']['WikiRAG'] - up['base_seconds']['WikiRAG']:+.3f}s）",
         "分钟级（瓶颈是条目重写；编码部分 0.01s）"),
    ]
    for i, vals in enumerate(eng, start=1):
        for cell, v in zip(t6.rows[i].cells, vals):
            fill_cell(cell, v)
    insert_after_table(t6, NOTE_44[0], p_cap6)

    # ---- 4.5 MRR 手算
    replace_block(doc, "提示：任选 3 题（须含至少一道多跳题）", NOTE_45)
    t7 = doc.tables[9]
    mrr = [("Vector RAG", "Q01 / Q08 / Q11", "1、1、1", "1/1 + 1/1 + 1/1 = 3.000 → 3.000/3 = 1.000",
            "一致（--mrr_check 1,1,1 → 1.000）"),
           ("GraphRAG", "Q01 / Q08 / Q11", "1、1、未命中(0)", "1/1 + 1/1 + 0 = 2.000 → 2.000/3 = 0.667",
            "一致（--mrr_check 1,1,0 → 0.667）"),
           ("WikiRAG", "Q01 / Q08 / Q11", "1、2、5", "1/1 + 1/2 + 1/5 = 1.700 → 1.700/3 = 0.567",
            "一致（--mrr_check 1,2,5 → 0.567）")]
    for i, vals in enumerate(mrr, start=1):
        for cell, v in zip(t7.rows[i].cells, vals):
            fill_cell(cell, v)
    add_row(t7, ("（教材手算例）", "—", "3、1、2", "1/3 + 1/1 + 1/2 = 1.833 → 1.833/3 = 0.611",
                 "一致（--mrr_check 3,1,2 → 0.611）"))

    # ---- 5.1 分析
    replace_block(doc, "提示：结合表 3 与表 4 回答", ANALYSIS_51)

    # ---- 5.2 失败案例
    replace_block(doc, "提示：原因归类建议用标签", FAIL_NOTE)
    t8 = doc.tables[10]
    for i, vals in enumerate(FAIL_ROWS, start=1):
        for cell, v in zip(t8.rows[i].cells, vals):
            fill_cell(cell, v)
    add_figure_after(doc, t8._tbl, os.path.join(IMG, "fig6_graph_q11.png"), 13.5,
                     "图 6  GraphRAG 在全局题 Q11 上的 local / global 两种查询对比"
                     "（local 只命中 doc01、Recall=0；global 命中 C2 社区、3/3 全中）")

    # ---- 6 AI 协作记录
    p6 = next(p for p in doc.paragraphs if p.text.strip().startswith("6  AI 协作记录"))
    set_par_text(p6, "6  AI 协作记录（编码智能体使用情况）")
    insert_after(p6, "说明：本实验的编码与调试使用 DeepSeek Harness（DSH）编码智能体完成，"
                     "它在本实验中承担指南里 TRAE 的角色（含 Builder / Chat / 补全 / 调试/终端五类用法）。"
                     "下表如实记录本轮协作的提示词、AI 输出与我的采纳/否决。", template_par=p6)
    for p in doc.paragraphs:
        if p.text.strip().startswith("表 9  TRAE 协作过程记录"):
            set_par_text(p, "表 9  编码智能体（DSH）协作过程记录")
    t9 = doc.tables[11]
    while len(t9.rows) - 1 < len(AI_ROWS):
        add_row(t9, ("", "", "", ""))
    for i, vals in enumerate(AI_ROWS, start=1):
        for cell, v in zip(t9.rows[i].cells, vals):
            fill_cell(cell, v)

    # ---- 7 思考题
    answers = {
        "(1)": [
            "Q01 的完整关系链（至少两条三元组）：①（李文瀚, 主讲, 大模型通识课）来自 doc06；"
            "②（李文瀚, 任职于, 人工智能学院）来自 doc04；③（人工智能学院, 院长, 王启明）来自 doc02。"
            "回答「《大模型通识课》授课教师所在学院的行政负责人是谁」，实际上要沿这条 3 跳链走完。",
            "Vector RAG 为什么难以把链上的文档都排进 Top-5：问题里只出现一个实体线索「《大模型通识课》」，"
            "链上的另外两个实体（李文瀚、人工智能学院）与问题几乎没有字面重合，承接关系的那篇 doc02"
            "整篇只说「人工智能学院成立于 2018 年，首任及现任院长为王启明教授」——既没有课程名，"
            "也没有「行政负责人」这个词，问句与桥接文档之间缺少可对齐的词面/语义锚点。"
            "实测 Q01 中各文档余弦分为 doc04 0.6834 > doc06 0.6028 > doc05 0.5893 > doc11 0.5524 > "
            "doc07 0.5456（Top-5 门槛）> doc10 0.5041 > doc02 0.5027（第 7 位），"
            "doc02 只差 0.04 就被窗口切掉——它是「被相似度排到窗口外」，而不是「库里没有」。"
            "更根本的原因是：单个查询向量只能表达问句的一个语义中心，「课程」这个中心与「院长」这个答案"
            "在向量空间里相距很远，一次相似度检索无法同时覆盖链条的两端。",
            "GraphRAG 靠什么机制沿链取全：它把检索从「整句相似度」换成「结构遍历」两步走——"
            "第一步实体链接：剥掉书名号后在图上精确命中节点「大模型通识课」（这一步不依赖整句相似度，"
            "只要有实体名即可）；第二步有界沿边展开：以该节点为起点做最多 3 跳的 BFS/DFS，"
            "每展开一个实体就把该实体所有关联边（连同边上的来源文档号）收进候选池，同时用「已展开集合」"
            "防止环上重复展开、按出度优先保证关系链先走通。于是 1 跳拿到 doc06、2 跳拿到 doc04/doc11、"
            "3 跳拿到 doc02，再按候选邻域文本与疑问句的相似度排序，Top-5 内即包含全部 3 篇金标准文档"
            "（Q01 Recall=1.0、MRR=1.0）。一句话总结：Vector RAG 是「向量的近邻」，GraphRAG 是「图的邻居」——"
            "前者靠相似度、后者靠拓扑，而多跳链的每一环恰恰是靠拓扑连起来的。"
            "代价也要说清楚：这套机制的前提是实体必须被抽进图且写法一致，Q16（人工智能分馆）与"
            "「云大/云山大学」这类同义说法就是它的两个真实软肋。",
        ],
        "(2)": [
            "MRR 计算：三题首次命中排名为 2、1、4 → MRR = (1/2 + 1/1 + 1/4) / 3 = (0.500 + 1.000 + 0.250) / 3 "
            "= 1.750 / 3 ≈ 0.583。",
            "两个指标分别在衡量什么：Recall@5（式 8-9）= 命中的金标准文档数 ÷ 金标准文档总数，"
            "衡量「该找的都找齐了吗」，是覆盖度指标，只关心证据有没有落进窗口，不关心顺序；"
            "MRR（式 8-10）= 各题首个命中文档排名倒数的平均，衡量「第一个正确结果排得多靠前」，"
            "是排序质量指标，只关心第一条命中的位置，不关心其余证据是否齐全。",
            "为什么必须一起看：Recall@5 全为 1.0 而 MRR 仅 0.4，意味着三题证据都在窗口内，"
            "但首命中排名是 1、5、5 这类分布——从「生成」的角度看，拼上下文时窗口前部会被无关内容占据，"
            "LLM 很容易被带偏或被淹没，检索的窗口预算被浪费了；只看 Recall 会以为一切正常。"
            "反过来只看 MRR 也不够：本实验 GraphRAG 在 Q11（scope=local）上 Recall=0、MRR=0，"
            "两项都低才是「根本没检到」的真失败，而 WikiRAG 在 Q08 上是 Recall=1.0、MRR=0.5，"
            "属于「检到了但排太后」的假失败——两类问题的修法完全不同：前者要扩大召回（加路径、加 K），"
            "后者要改进排序（重排、融合）。本实验的 Hybrid RAG 正是针对后者："
            "它把宏平均 Recall@5 从 0.931 提到 0.950，MRR 却从 0.948 降到 0.921，恰好说明两个指标在牵着不同的方向走。",
        ],
        "(3)": [
            "Vector RAG 会怎么失败：Q11「云山大学有哪些科研平台？分别依托谁建设？」的答案由两半组成——"
            "「认知计算全国重点实验室 ↔ 人工智能学院」（doc10、doc02）与「岭南超算中心 ↔ 计算机学院」（doc10、doc03）。"
            "系统把问题编码成一个向量，最像的是集中讲平台的 doc10（排第 1），但另外两半证据在别的文档里、"
            "与这条问句的整体语义相似度更低，于是落在窗口之外：实测 Q11 检索到 ['doc10','doc05','doc15','doc01','doc03']，"
            "漏掉 doc02，Recall 只有 0.667。它失败的本质是「一个查询向量表达不了多个子问题」。",
            "GraphRAG 的社区摘要为什么适合：全局问题的答案往往不是某一句，而是「一群实体之间的整体格局」。"
            "GraphRAG 在建图之后先用社区发现把稠密子图聚成社区，再为每个社区写一段自足的摘要（教材 8.7 节），"
            "于是检索单位从「文档」变成「社区」，一次命中就带回该社区的全部来源文档。"
            "实测 Q11 用 scope=global 命中 C2（相似度 0.5703），其来源文档恰好是 doc02/doc03/doc10/doc15，"
            "一次就拿到 3/3、首次命中排名 1（对比同一题 scope=local 只有 doc01、Recall=0）。",
            "WikiRAG 又靠什么机制解决：它把「整理知识」这一步前置到了索引期——人工/LLM 把散落的原文重写成"
            "「一个条目一个主题」的聚合正文，检索单位变成条目。例如「人工智能学院」条目里同时写了"
            "「认知计算全国重点实验室依托本院建设」和「与新加坡南洋理工大学有本科联合培养项目」，"
            "命中条目即等同于同时取回 doc02/doc10/doc15。实测 Q20（地理位置+交通+校历）WikiRAG 是唯一拿到 3/3 的范式，"
            "因为「云山大学」条目已经把 doc01/doc12/doc14 聚合在一起了。",
            "三者的本质差异：Vector RAG 检索的是「文本片段的语义相似度」，知识组织最轻、更新最便宜，"
            "但检索单位（块）与「答案的完整性」没有必然联系；GraphRAG 检索的是「关系的结构」，"
            "用拓扑把跨文档的关系链一次取全，代价是抽取质量与更新成本；"
            "WikiRAG 检索的是「被重写过的知识条目」，把人的领域判断固化进索引，"
            "在全局归纳上最稳，代价是重写成本与细节损失。用一句话概括："
            "Vector 靠相似度、GraphRAG 靠结构、WikiRAG 靠组织——前者决定「像不像」，"
            "中者决定「连不连得上」，后者决定「整理得好不好」。",
        ],
        "(4)": [
            "三种范式各自要改什么（以「新增一位教师，他主讲一门新课」为例）：",
            "Vector RAG：把新教师写进文档后重新切块，只重新编码受影响的新块（若写进已有文档，则该文档的所有块"
            "都要重切重编码），再往向量索引追加。实测把 doc16 加入语料后，全量重建只从 0.344s 变成 0.361s"
            "（+0.017s）——几乎没有成本，而且不需要任何 LLM 调用。",
            "GraphRAG：先用 LLM 对新文档做三元组抽取（实测 15 篇 81 条耗时 12.8s，约 0.85s/篇），"
            "把新实体并入图（本实验 45→48 实体、40→43 三元组），重跑社区发现（社区数不变，仍是 10），"
            "并重新生成受影响社区的摘要（这一步是 LLM 成本的大头）。自动部分实测 0.113s，但抽取与摘要的"
            "LLM 调用是它独有的、也是最大的一笔开销。",
            "WikiRAG：先由人（或 LLM）判断新教师是并入既有的「人工智能学院教师」条目，还是新建一个条目，"
            "然后重写该条目正文并重新编码（实测 +0.013s）。编码几乎免费，真正的成本是「重写」这一步的判断与写作。",
            "从更新成本看表 8-4 的工程定位：Vector RAG 是「低成本、可无限扩展」的默认基线——"
            "语料变成几万篇也只是多切些块、多编码些向量，机器可以全自动完成，"
            "所以它适合做第一层召回与大规模语料的兜底；GraphRAG 是「高成本、高价值」的关系型知识层——"
            "每次更新都涉及 LLM 抽取与摘要重写，但它换来的是多跳与全局能力（本实验多跳 Recall 1.000、"
            "全局题用 Global 查询从 0.639 提到 0.750），适合变化不频繁、但对关系推理要求高的知识库"
            "（如规章制度、组织架构、科研关系网）；WikiRAG 是「中等成本、面向主题」的知识产品——"
            "更新时要人做判断，但一旦组织好，它在极小窗口（K=1~2）和全局题上最稳，"
            "适合把高频主题沉淀成稳定的知识条目（FAQ、词条库）。工程上更常见的做法是三层并用："
            "Vector 做召回、Graph 做关系补全、Wiki 做高频主题的确定性回答。",
        ],
        "(5)": [
            "忠实度人工评分主观性的缓解办法：① 先写死评分细则再打分——本实验采用「逐句核查」："
            "5 分＝每句都能在检索到的资料里找到依据，3 分＝主结论有依据但存在超出资料的推断，"
            "1 分＝主要内容无依据；② 评分时必须同时保留「答案 + 检索到的资料原文」作为证据"
            "（本实验把 15 组答案与上下文一起导出到 out/faith.json，任何一条分数都能回溯核对）；"
            "③ 记录「无依据的句子」原文而不是只给总分，让分歧可定位；④ 多题、多范式交叉打分，"
            "避免因单题印象而整体偏高或偏低；⑤ 必要时换人或两人独立打分取平均（Kappa 一致性）。",
            "RAGAS 用 LLM 自动评忠实度的思路（式 8-8）：把答案拆成若干条原子断言（claim），"
            "对每条断言让 LLM 判断「能否由给定的检索上下文推出」，"
            "忠实度 = 被上下文支持的断言数 ÷ 断言总数。这样把「人工逐句核查」这一步自动化，"
            "既省人力又能给出连续的分数（而不是通过/不通过），并且可以与检索指标一起做批量回归测试。",
            "「用 LLM 评 LLM」引入的新偏差：① 自我偏好——评审模型倾向于给与自己生成风格相似的答案更高分，"
            "尤其是同一模型既生成又评审时；② 知识泄漏——对「事实上正确但资料里没写」的断言，"
            "评审模型可能凭自己的先验判为「支持」，从而高估忠实度；③ 位置与长度偏差——"
            "上下文的顺序、答案的长短会影响判定；④ 粒度问题——断言怎么切分本身没有唯一标准，"
            "切得粗会掩盖局部幻觉，切得细会把同义改写误判为无依据；"
            "⑤ 上限受评审模型能力约束，且系统性错误会被同源模型复制。"
            "本实验的一个对照很说明问题：在「禁止编造」的严格 prompt 下，15 组答案全部得 5 分；"
            "而把约束放宽（不许拒答）后，同样的问题出现了无依据的句子——Q20 的「通常按秋季学期 9 月开学、"
            "春季学期 2 月或 3 月开学」（检索结果里根本没有 doc12，且与金标准部分冲突），只能拿 2 分。"
            "这说明忠实度分数同时受「检索是否给全资料」和「模型是否守规矩」两个因素影响，"
            "单看分数无法区分二者，必须与 Recall/MRR 联合解读——这也正是本实验把三类指标放在一起测的原因。",
        ],
        "(6)": [
            "协作对象说明：本实验使用 DeepSeek Harness（DSH）编码智能体，承担指南中 TRAE 的角色。"
            "抽取 prompt 一共迭代了 2 版：v1 是指南给的模板（限定「只输出 JSON 数组、关系用 2~4 字短语、"
            "一篇文档 5~10 条」）；v2 追加了「不要 Markdown 代码块、不要解释文字，只输出一个 JSON 数组」的"
            "指令并作为失败重试的兜底。实测 15 篇文档一次通过、抽出 81 条三元组、首次解析失败率 0%，"
            "v2 兜底分支一次都没被触发——这个结果本身说明：把格式约束写进 prompt 比在解析端做容错更省事。",
            "最关键的一次改动是「实体归一化」：原来的匹配是「节点名是否是问题的子串」，"
            "于是问题里的「《大模型通识课》」永远匹配不到图节点「大模型通识课」（多了书名号），"
            "「云大」也匹配不到「云山大学」。改成 ENTITY_ALIASES 别名表 + normalize_text() 剥离书名号/引号/空白之后，"
            "含书名号的课程类问题不再退化为社区全局查询——这是 GraphRAG 能不能在多跳题上发挥作用的开关。",
            "AI 生成的代码里我必须亲自检查的三处：① JSON 解析——模型输出可能带前后缀说明或代码块，"
            "必须用 raw.index(\"[\")/raw.rindex(\"]\") 截取，并统计解析失败率（指南 Q3 要求报告这个数字）；"
            "② 实体归一化——别名表合并错误会把不同实体并成一个节点（例如「图书馆」与「人工智能分馆」"
            "在本实验里必须保持两个节点，否则 Q16 的失败原因就查不清了）；"
            "③ 多跳展开的边界——必须有「已展开集合」防止有环图反复展开同一实体，同时限制跳数与每跳邻域上限，"
            "否则一次查询会把整张图卷进候选池，既慢又稀释排序。",
            "另外记一次被否决的 AI 建议：AI 提出把 GraphRAG 全局查询的「社区文档拼接后截断」改成"
            "「多社区轮转交错」，理由是这样可以避免第一个社区占满 Top-K。我实测后发现全局题 Recall 从 0.750"
            "降到 0.583、MRR 从 0.506 降到 0.458——因为全局题的证据本来就高度集中在相似度最高的那个社区，"
            "拓宽窗口反而引入噪声。这个负结果连同「社区数 1→4 对结果完全没有影响」的扫描数据一起写进了报告 5.1 节。"
            "这件事给我的教训是：AI 给的「听起来合理」的优化必须用数据验一遍，"
            "而验证的前提是评测脚本本身可靠（指标口径统一映射回文档级、K 扫描逐档重跑而不是截断）。",
        ],
    }
    for key, lines in answers.items():
        paras = list(doc.paragraphs)
        idx = next(i for i, p in enumerate(paras) if p.text.strip().startswith(key))
        q_par = paras[idx]
        j = idx + 1
        holders = []
        while j < len(paras) and is_placeholder(paras[j]):
            holders.append(paras[j])
            j += 1
        anchor = holders[0] if holders else q_par
        set_par_text(anchor, lines[0])
        last = anchor
        for line in lines[1:]:
            last = insert_after(last, line, template_par=anchor)
        for h in holders[1:]:
            h._p.getparent().remove(h._p)

    # ---- 附录 A：20 题逐题明细
    p_end = doc.paragraphs[-1]
    p_end.add_run()
    tbl = doc.add_table(rows=1, cols=6)
    tbl.style = "Table Grid"
    hdr = ("题号", "类型", "金标准 evidence", "Vector RAG 首命中排名", "GraphRAG 首命中排名", "WikiRAG 首命中排名")
    for cell, v in zip(tbl.rows[0].cells, hdr):
        fill_cell(cell, v)
    detail = {r["id"]: r for r in data["detail.json"]["rows"]}
    for q in H.QUESTIONS:
        r = detail[q["id"]]
        add_row(tbl, (q["id"], H.TYPE_CN[q["type"]] + ("（自拟）" if q["src"] == "自拟" else ""),
                      "、".join(q["evidence"]),
                      str(r["Vector RAG"]["first_rank"] or "未命中"),
                      str(r["GraphRAG"]["first_rank"] or "未命中"),
                      str(r["WikiRAG"]["first_rank"] or "未命中")))
    # 附录标题与说明放在表格前面
    cap = doc.add_paragraph("附录 A  20 题逐题首次命中排名（K=5，GraphRAG scope=local）")
    tbl._tbl.addprevious(cap._p)
    for line in reversed(APPENDIX_HEAD):
        ph = doc.add_paragraph(line)
        cap._p.addprevious(ph._p)

    # ---- 追加附件说明
    doc.add_paragraph("附件说明：① hw4_rag.py —— 本实验主程序（含 [修改1]~[修改7] 标注）；"
                      "② out/ 目录 —— 全部实验原始输出（eval.json、scan.json、detail.json、extract.json、"
                      "faith.json、faith_loose.json、update.json 及对应运行日志）；"
                      "③ report_img/ —— 报告插图源文件；④ make_figures.py、run_sweeps.py、"
                      "run_sweeps2.py、fill_report.py —— 配图、批量实验与本报告的生成脚本。")

    doc.save(OUT)
    print("[ok] 已生成", OUT)


if __name__ == "__main__":
    main()
