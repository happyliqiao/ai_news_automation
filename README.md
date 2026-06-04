# AI 新闻自动化

这个目录包含一个 Windows 自动化程序，用来每天生成两份 Word 日报：

- `AI相关新闻日报_yyyyMMdd_HHmmss.docx`
- `大杂烩新闻日报_yyyyMMdd_HHmmss.docx`

两份报告都会自动抓取 RSS/新闻搜索源、去重、按时效和关键词排序、尽量抓取正文，并生成更长的中文总结、关键要点、来源时间和原文链接。

## 报告内容

AI 相关新闻日报覆盖：

- AI、人工智能、大模型、OpenAI、ChatGPT、Claude、DeepSeek、Gemini、NVIDIA、AI Agent、AI 监管、安全、开源、融资等。

大杂烩新闻日报覆盖：

- 政治、财经、体育、娱乐、教育、生活、国际、国内、科技、汽车、游戏、育儿、职场等。

## 输出位置

- Word 文档：`E:\AI\yyyy-MM-dd\`
- 运行日志：`E:\AI\logs\`

程序会按日期自动创建文件夹，例如：

```text
E:\AI\2026-05-29\
```

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

## 安装每天 9:00 的计划任务

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install_daily_task.ps1
```

## 查看计划任务

```powershell
Get-ScheduledTask -TaskName "AI News Daily Report"
```
