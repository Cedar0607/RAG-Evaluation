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

- `target_count`: 最终题数，默认 30。
- `candidate_multiplier`: 候选题倍率，默认 1.2，所以候选题大约 36 道。
- `max_candidates_total`: 候选题硬上限，默认 45，即使 LLM 多输出也会截断。
- `max_chunks_per_doc`: 每篇文档最多处理的 chunk 数，默认 8。
- `max_total_chunks`: 本次运行最多处理的总 chunk 数，默认 40。
- `knowledge_points_per_chunk`: 每个 chunk 最多抽取的知识点数，默认 6。
- `max_consecutive_llm_errors`: 连续 LLM 失败阈值，默认 3，达到后停止后续 LLM 调用并保留已有结果。

如果只想更快地看一版质量，可以把 `max_chunks_per_doc` 改成 3~5，把 `target_count` 改成 10~20。

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
- `knowledge_points`：从文档抽取的知识点。
- `chunks`：Markdown 分块信息。
- `errors`：LLM 超时、解析失败等错误记录。
- `run_config`：本次运行配置摘要。

每道题会保留来源文档、章节、chunk id、证据摘要和 required evidence，后续可以接 Dify API 生成检索 chunk 和 answer，再用于 Ragas。
