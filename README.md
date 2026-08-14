# AI 新闻自动化

这个目录包含一个 Windows 自动化程序，用来每天生成两份 Word 日报：

- `AI相关新闻日报_yyyy-MM-dd.docx`
- `大杂烩新闻日报_yyyy-MM-dd.docx`

两份报告都会自动抓取 RSS/新闻搜索源、去重、按时效和关键词排序、尽量抓取正文，并生成中文总结、关键要点、来源时间和原文链接。可选功能：LLM 高质量中文摘要、微信/邮件推送。

## 报告内容

AI 相关新闻日报覆盖：

- AI、人工智能、大模型、OpenAI、ChatGPT、Claude、DeepSeek、Gemini、NVIDIA、AI Agent、AI 监管、安全、开源、融资等。

大杂烩新闻日报覆盖：

- 政治、财经、体育、娱乐、教育、生活、国际、国内、科技、汽车、游戏、育儿、职场等。

## 输出位置

- Word 文档：`output\yyyy-MM-dd\`（位于项目目录内，同一天多次运行会覆盖当天文件）
- 最新版固定入口：`output\最新\`（始终指向最近一次生成的文件，方便一键打开）
- 运行日志：`output\logs\`（位于项目目录内）

程序会按日期自动创建文件夹，例如：

```text
output\2026-05-29\
output\最新\
```

## 配置（可选）

复制 `.env.example` 为 `.env` 并填写（`.env` 已被 git 忽略，不会泄露密钥）。

### LLM 高质量中文摘要（推荐）

不配置也能运行，程序会自动使用本地抽取式摘要；配置后标题翻译、中文总结、关键要点质量显著提升：

```ini
AI_NEWS_API_KEY=sk-你的密钥
AI_NEWS_API_BASE=https://api.deepseek.com
AI_NEWS_MODEL=deepseek-chat
```

- 默认兼容 OpenAI 格式的接口（DeepSeek / OpenAI / 其他中转均可）
- LLM 调用失败会自动回退到本地摘要，不会中断生成

### 微信推送（任选其一）

```ini
# Server酱：https://sct.ftqq.com/ 微信扫码登录获取 SendKey
AI_NEWS_SERVERCHAN_KEY=SCTxxxx

# 或企业微信群机器人 Webhook
AI_NEWS_WECOM_WEBHOOK=https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxxx
```

### 邮件推送（可选）

```ini
AI_NEWS_SMTP_HOST=smtp.qq.com
AI_NEWS_SMTP_PORT=465
AI_NEWS_SMTP_USER=you@qq.com
AI_NEWS_SMTP_PASS=授权码
AI_NEWS_SMTP_TO=receiver@example.com
```

邮件会附带当天的 `.docx` 附件。推送失败只记录日志，不影响报告生成。

## 手动运行

在项目目录下执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\run_ai_news.ps1
```

只生成 AI 日报：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\run_ai_news.ps1 -Report ai
```

只生成大杂烩日报：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\run_ai_news.ps1 -Report general
```

直接运行 Python（可用参数：`--report`、`--limit`、`--detail-limit`、`--output-dir`、`--no-llm`、`--no-push`、`--api-key` 等，`python ai_news_daily.py --help` 查看全部）：

```powershell
python ai_news_daily.py
```

## 安装每天 9:00 的计划任务

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install_daily_task.ps1
```

## 查看计划任务

```powershell
Get-ScheduledTask -TaskName "AI News Daily Report"
```
