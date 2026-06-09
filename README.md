# 软件研发知识库 TestBench 出题脚本

这个脚本用于从 OCR 后的 Markdown 软件设计说明书中生成 RAG 评测用问题和标准答案，并输出 Excel，方便人工复核。

## 安装依赖

```powershell
pip install -r requirements.txt
```

## 配置

复制配置模板：

```powershell
Copy-Item config.example.json config.json
```

修改 `config.json` 中的本地 LLM 接口：

```json
{
  "vlm": {
    "endpoint": "http://localhost:8000/v1/chat/completions",
    "api_key": "your_api_key",
    "model": "Deepseek-V4-flash",
    "timeout": 120,
    "retries": 2
  }
}
```

## 运行

```powershell
python scripts/generate_questions.py --input-dir "D:\docs\markdown" --config config.json --output "output\questions_ground_truth.xlsx"
```

也可以指定多个文件：

```powershell
python scripts/generate_questions.py --input-files "D:\a.md" "D:\b.md" --config config.json --output "output\questions_ground_truth.xlsx"
```

## 快速试跑策略

默认配置以“快速生成几十个问题，方便人工检查”为目标：

- `generation_mode`: 默认 `section_detailed`，按章节直接生成综合型问题和较长 Ground Truth。
- `target_count`: 最终题数，默认 30。
- `candidate_multiplier`: 候选题倍率，默认 1.2，所以候选题大约 36 道。
- `max_candidates_total`: 候选题硬上限，默认 45，即使 LLM 多输出也会截断。
- `max_chunk_chars`: 单个章节 chunk 最大字符数，默认 80000。
- `max_chunks_per_doc`: 每篇文档最多处理的高价值章节 chunk 数，默认 5。
- `max_total_chunks`: 本次运行最多处理的总 chunk 数，默认 30。
- `questions_per_chunk`: 每个章节 chunk 最多生成的问题数，默认 2。
- `answer_min_paragraphs` / `answer_max_paragraphs`: Ground Truth 默认 2~4 个自然段。
- `temperature`: 默认 0.1，降低随机性和术语自由发挥。
- `max_consecutive_llm_errors`: 连续 LLM 失败阈值，默认 3，达到后停止后续 LLM 调用并保留已有结果。

如果只想更快地看一版质量，可以把 `max_chunks_per_doc` 改成 3~5，把 `target_count` 改成 10~20。

## 当前切分方式

脚本会先按 Markdown 标题切分章节。如果某个章节超过 `max_chunk_chars`，才会按长度继续切分，并使用 `chunk_overlap_chars` 做少量重叠。

在默认的 `section_detailed` 模式下，脚本不会均匀抽样 chunk，而是会跳过目录、修订记录、术语、参考文献等低价值章节，再根据标题和内容中的关键词打分，优先选择功能、设计、架构、流程、接口、模块、软件层、CPD、异常、状态、数据、时序、交互、约束等章节。

选中的每个章节会直接输入 LLM 生成综合问答，Ground Truth 会要求写成多段解释，适合软件设计文档中的 CPD 项功能说明、软件层架构说明、模块协作和端到端流程。

提示词会要求缩写、英文术语、变量名、模块名、接口名和专有名词保持原文写法。除非来源章节明确给出了定义，否则不得自行翻译、展开缩写或添加括号释义，例如不得擅自把 `operator` 写成 `operator（操作员）`。

## 失败与中途保存

脚本默认开启 `continue_on_error` 和 `checkpoint_enabled`：

- 某次 LLM 超时或输出格式异常时，会记录到 `errors` sheet，并跳过当前 chunk/文档继续运行。
- 每处理完一个 chunk 会覆盖写入一次 Excel。
- 同时会写入 `output\questions_ground_truth.checkpoint.json`，用于保留结构化中间结果。
- 如果最终 LLM 终筛失败，脚本会用候选题兜底生成 `final_questions`，避免空结果。

## Excel 输出

生成的 Excel 包含：

- `final_questions`：最终题集，约 30 题。
- `candidates`：候选题，方便追溯被筛掉的题。
- `knowledge_points`：旧的知识点模式会使用；默认 `section_detailed` 模式下通常为空。
- `chunks`：Markdown 分块信息。
- `errors`：LLM 超时、解析失败等错误记录。
- `run_config`：本次运行配置摘要。

每道题会保留来源文档、章节、chunk id、证据摘要和 required evidence，后续可以接 Dify API 生成检索 chunk 和 answer，再用于 Ragas。

## Ragas 两题试跑

`ragas_test.py` 读取包含以下四列的 Excel：

- `question`
- `answer`
- `contexts`
- `ground_truths`

`contexts` 推荐保存为 JSON 数组字符串，例如：

```json
["父段内容1", "父段内容2"]
```

Ragas 0.4.3 需要 Python 3.10 或更高版本。安装依赖并创建配置：

```powershell
pip install -r requirements.txt
Copy-Item ragas_config.example.json ragas_config.json
```

修改 `ragas_config.json` 中的 Judge LLM 接口后运行：

```powershell
python ragas_test.py --input "D:\data\ragas_input.xlsx" --output "output\ragas_result.xlsx" --limit 2
```

脚本会评估 `context_precision`、`context_recall`、`faithfulness` 和 `factual_correctness`，把逐题结果写入原数据 sheet，并创建 `ragas_summary` 汇总 sheet。每处理完一题都会保存输出文件。

Judge LLM 仍通过 `ragas_config.json` 中的 OpenAI-compatible 本地接口调用。脚本包含针对 Ragas 0.4.3 旧 VertexAI 可选导入路径的兼容处理，不需要为本地 LLM 安装或配置 VertexAI。
