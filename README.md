# ModelTest

面向 OpenAI Chat Completions 兼容接口的模型功能、长上下文、输出容量和压力测试工具。工具使用 Python 标准库，可保存逐请求证据并生成 HTML、CSV 和 JSON 报告。

## 功能

- 接口连通性与基础生成测试
- 连续序号、重复输出和长文写作测试
- DOCX/TXT 文档问答、摘要、审核与翻译
- 文档差异分析
- 输入、输出和总上下文边界测试
- 可配置并发压力测试
- 保存请求、响应、正文、推理字段、usage 和服务端扩展指标

## 环境

- Python 3.9 或更高版本
- 可访问的 OpenAI Chat Completions 兼容接口
- 无第三方 Python 依赖

## 配置

编辑 `automation/config.json`：

```json
{
  "endpoint": "http://127.0.0.1:1025/v1/chat/completions",
  "models": ["deepseek14b"],
  "stream": false,
  "max_tokens": 4096
}
```

如果接口需要令牌，在运行前设置环境变量。令牌不会写入请求日志：

```powershell
$env:MODEL_API_KEY = "your-token"
```

```bash
export MODEL_API_KEY="your-token"
```

请根据实际服务配置同步调整 `server_limits`、token 档位和超时。`max_seq_len` 表示输入与输出的总长度上限，不应直接作为所有请求的 `max_tokens`。

## 测试文档

脚本从仓库根目录读取 `.docx`、`.doc`、相关 `.txt` 以及历史 `.xlsx` 测试资料。`docs/` 只保存测试方法、用例和优化说明，不作为业务文档输入目录。业务文档默认被 `.gitignore` 排除，不会意外提交到仓库。

没有文档时仍可运行不依赖文档的用例；相关文档用例会按程序规则跳过或减少覆盖范围。

当前目录结构：

```text
modeltest/
  automation/       测试程序与配置
  docs/             测试说明和用例文档
  模型测评结果/       自动生成的报告与逐请求证据
  Linux测试结果/     Linux 历史场景执行结果
  *.doc/*.docx      本地业务测试资料（Git 忽略）
  *.txt/*.xlsx      文档转换文本和历史用例（Git 忽略）
```

详细文档入口见 [`docs/README.md`](docs/README.md)。

## 运行

```powershell
python .\automation\run_tests.py --dry-run
python .\automation\run_tests.py --mode smoke
python .\automation\run_tests.py --mode functional --model deepseek14b
python .\automation\run_tests.py --mode full --model deepseek14b
python .\automation\run_tests.py --mode capacity --model deepseek14b
python .\automation\run_tests.py --mode limits --model deepseek14b
python .\automation\run_tests.py --mode stress --model deepseek14b
```

Linux 使用相同参数：

```bash
python3 automation/run_tests.py --mode smoke --model deepseek14b
```

可用模式：

| 模式 | 用途 |
|---|---|
| `smoke` | 连通性和少量基础用例 |
| `functional` | 功能、写作和文档处理 |
| `capacity` | 输出 token 容量测试 |
| `limits` | 输入及上下文边界测试 |
| `stress` | 配置的并发压力测试 |
| `extended` | 扩展并发档位 |
| `full` / `all` | 完整测试 |

## 输出

结果生成在仓库根目录的 `模型测评结果/`：

```text
模型测评结果/
  report.html
  results.csv
  results.json
  <timestamp>/
    report.html
    run_config.json
    cases.json
    requests/
```

状态说明：

- `OK`：接口正常结束并返回正文
- `TRUNCATED`：达到输出 token 上限
- `EXPECTED_TRUNCATED`：测试设计允许触达输出上限
- `ERROR` / `TIMEOUT`：请求错误或超时
- `SKIPPED`：资料或前置条件不足
- `PASS` / `FAIL`：确定性自动规则结果
- `REVIEW`：需要人工评价内容质量

## 本地回归测试

```powershell
python .\automation\test_runner.py
```

`test_runner.py` 使用本地模拟 HTTP 服务，不需要连接真实模型。

## 安全说明

- 不要把 API 密钥写入 `config.json`。
- 不要提交真实业务文档、生成结果或服务日志。
- 发布前检查 endpoint、请求证据和文档内容是否包含内部地址或敏感信息。
