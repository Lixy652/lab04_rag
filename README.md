# 实验作业四 · Vector RAG / GraphRAG / WikiRAG 对比实验

> 提交物：`<学号>_<姓名>_实验作业四.docx`（实验报告，本地生成、未随仓库分发）+ `hw4_rag.py`（含本人修改的程序）

## 一、目录说明

| 路径 | 说明 |
|---|---|
| `hw4_rag.py` | **主程序**（单文件可运行）。以指南第七节参考框架为底稿，改动处均以 `[修改1]`~`[修改7]` 标注 |
| `make_figures.py` | 配图生成脚本（检索命中截图、知识图谱、条目库、K 敏感性曲线等） |
| `run_sweeps.py` / `run_sweeps2.py` | 批量实验驱动（一次加载模型跑完全部评测/消融/扫描） |
| `fill_report.py` | 按配套样板生成实验报告的脚本（学号/姓名/班级由 `--sid/--name/--class` 传入，不写死在源码里） |
| `out/`、`report_img/` | 脚本生成的实验结果与插图（默认已被 `.gitignore` 排除，可由上面的脚本重新生成） |
| `sessionfiles/`、`guide.txt`、`template.txt` | 课程材料（指南 PDF、报告样板及其提取文本），**未随仓库分发** |

> `.gitignore` 已排除：实验报告 docx（含学号姓名）、与作业无关的本地记录、课程材料、以及可重新生成的 `out/` 与 `report_img/`。
> 若确实想把实验结果一并发布，删掉 `.gitignore` 中 `out/`、`report_img/` 两行即可。

## 二、复现命令

```powershell
pip install sentence-transformers networkx numpy openai          # 依赖
$env:HF_ENDPOINT = "https://hf-mirror.com"                        # 模型下载镜像（首次）
$env:LLM_API_BASE="https://api.deepseek.com"; $env:LLM_MODEL="deepseek-chat"
$env:LLM_API_KEY="<你的 key>"                                     # 仅环境变量注入，不写入任何文件

python hw4_rag.py --mode eval                                     # 三范式 × 20 题检索评测
python hw4_rag.py --mode scan                                     # K=1/2/3/5 敏感性扫描
python hw4_rag.py --mode detail --out out/detail.json             # 逐题明细
python hw4_rag.py --mode graph --q Q11 --scope local / global     # Local vs Global 对比
python hw4_rag.py --mode faith --q Q01,Q06,Q11 --out out/faith.json
python hw4_rag.py --mode extract                                  # LLM 现场抽三元组（需 API）
python hw4_rag.py --mode update                                   # 新增 doc16 的增量更新账
python hw4_rag.py --mrr_check 1,2,5                               # MRR 手算复核（式 8-10）
python run_sweeps.py; python run_sweeps2.py                       # 批量实验（含消融）
python make_figures.py; python fill_report.py                     # 生成配图与报告
```

## 三、主要实验结论（数据来自 `out/eval.json`、`out/scan.json`）

**20 题（15 内置 + 5 自拟）文档级评测**

| 范式 | 事实 Recall@5 / MRR | 多跳 Recall@5 / MRR | 全局 Recall@5 / MRR |
|---|---|---|---|
| Vector RAG | 1.000 / 1.000 | 0.905 / 0.929 | 0.889 / 0.917 |
| GraphRAG | 0.857 / 0.857 | **1.000** / 0.857 | 0.639 / 0.478 |
| WikiRAG | 1.000 / 0.786 | 0.905 / 0.821 | 0.722 / 0.589 |

1. **多跳题是 GraphRAG 的主场**：Q01 沿「大模型通识课 ←主讲— 李文瀚 —任职于→ 人工智能学院 —院长→ 王启明」三跳取全 3 篇证据（Recall 1.0、MRR 1.0）；Vector RAG 的桥接文档 doc02 余弦仅 0.5027、排第 7，被 Top-5 窗口切掉。
2. **全局题必须换查询方式**：Q11 用 `scope=local` 只命中 doc01（Recall 0），换 `scope=global` 命中 C2 社区后 3/3 全中——Local 与 Global 是互补能力。
3. **窗口预算（教材 8.3 节）**：K=1 时多跳题在结构上不可能满分（证据有 2~3 篇）；全局题对 K 最敏感（GraphRAG 0.700→0.200）。事实题几乎不受 K 影响（证据只有 1 篇）。
4. **忠实度**：「禁止编造」约束下 15 组答案全部 5 分，但低召回表现为「资料中未提及」而非幻觉；放宽约束的对照中 Q20 出现无依据句子（2 分），说明忠实度必须与 Recall/MRR 联合解读。
5. **增量更新**：三范式自动重建都是亚秒级（Vector +0.017s、GraphRAG ≈0、Wiki +0.013s），真正的成本在 LLM 抽取（实测 0.85s/篇）与条目重写。
6. **负结果（如实记录）**：社区数 1→4 对结果无影响（拼接后按 K 截断，第一个社区就占满窗口）；把拼接改成轮转交错反而使全局题 Recall 从 0.750 降到 0.583 —— 已写入报告 5.1 节。

## 四、隐私与版权说明

- 实验报告成品（`<学号>_<姓名>_实验作业四.docx`，封面含学号/姓名）与课程材料（`sessionfiles/`、`guide.txt`、`template.txt`）均已被 `.gitignore` 排除，**不随仓库发布**；报告正文中只写「密钥通过环境变量 `LLM_API_KEY` 注入」，不含任何密钥。
- 生成报告用的 `fill_report.py` 不再写死学号/姓名/班级，改由 `--sid/--name/--class`（或同名环境变量）传入，仓库源码中不残留个人信息。
- `out/` 与 `report_img/` 默认也不提交：前者是运行日志（含本机绝对路径），两者都可由 `run_sweeps*.py` / `make_figures.py` 重新生成。
