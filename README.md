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

## Excel 输出

生成的 Excel 包含：

- `final_questions`：最终题集，约 30 题。
- `candidates`：候选题，方便追溯被筛掉的题。
- `knowledge_points`：从文档抽取的知识点。
- `chunks`：Markdown 分块信息。
- `run_config`：本次运行配置摘要。

每道题会保留来源文档、章节、chunk id、证据摘要和 required evidence，后续可以接 Dify API 生成检索 chunk 和 answer，再用于 Ragas。
