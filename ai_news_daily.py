#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Daily news collector.

Generates two Word reports:
- AI相关新闻日报
- 大杂烩新闻日报

Reports are saved under E:\\AI\\yyyy-MM-dd by default.
"""

from __future__ import annotations

import argparse
import html
import json
import re
import textwrap
import time
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


OUTPUT_DIR = Path(r"E:\AI")
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36"
)


@dataclass(frozen=True)
class ReportConfig:
    key: str
    title: str
    filename_prefix: str
    search_terms: list[str]
    keywords: list[str]
    extra_feeds: list[str]
    categories: dict[str, list[str]]


@dataclass(frozen=True)
class NewsItem:
    title: str
    link: str
    source: str
    published: datetime | None
    summary: str
    category: str


@dataclass(frozen=True)
class EnrichedNewsItem:
    item: NewsItem
    article_text: str
    chinese_title: str
    chinese_summary: str
    key_points: list[str]


AI_REPORT = ReportConfig(
    key="ai",
    title="AI相关新闻日报",
    filename_prefix="AI相关新闻日报",
    search_terms=[],
    keywords=[
        "AI",
        "人工智能",
        "大模型",
        "模型",
        "OpenAI",
        "ChatGPT",
        "GPT",
        "Claude",
        "Anthropic",
        "DeepSeek",
        "Gemini",
        "智能体",
        "Agent",
        "算力",
        "芯片",
        "NVIDIA",
        "监管",
        "安全",
        "开源",
        "融资",
        "发布",
    ],
    extra_feeds=[
        "https://openai.com/news/rss.xml",
        "https://www.anthropic.com/news/rss.xml",
        "https://ai.googleblog.com/feeds/posts/default",
        "https://blogs.nvidia.com/feed/",
        "https://huggingface.co/blog/feed.xml",
    ],
    categories={},
)


AI_FIXED_FEEDS: dict[str, list[str]] = {
    "AI": [
        "https://openai.com/news/rss.xml",
        "https://www.anthropic.com/news/rss.xml",
        "https://ai.googleblog.com/feeds/posts/default",
        "https://blogs.nvidia.com/feed/",
        "https://huggingface.co/blog/feed.xml",
    ],
    "科技": [
        "https://www.chinanews.com.cn/rss/it.xml",
        "https://www.chinanews.com.cn/rss/finance.xml",
    ],
}


GENERAL_REPORT = ReportConfig(
    key="general",
    title="大杂烩新闻日报",
    filename_prefix="大杂烩新闻日报",
    search_terms=[],
    keywords=[
        "政治",
        "财经",
        "经济",
        "股市",
        "体育",
        "娱乐",
        "教育",
        "生活",
        "国际",
        "国内",
        "科技",
        "汽车",
        "游戏",
        "育儿",
        "职场",
        "政策",
        "消费",
        "社会",
    ],
    extra_feeds=[],
    categories={
        "政治": ["政治 新闻", "政策 时政"],
        "财经": ["财经 新闻", "经济 股市 消费"],
        "体育": ["体育 新闻", "足球 篮球 体育"],
        "娱乐": ["娱乐 新闻", "电影 综艺 明星"],
        "教育": ["教育 新闻", "高考 学校 教育"],
        "生活": ["生活 新闻", "健康 消费 生活"],
        "国际": ["国际 新闻", "全球 国际"],
        "国内": ["国内 新闻", "中国 社会 新闻"],
        "科技": ["科技 新闻", "互联网 科技"],
        "汽车": ["汽车 新闻", "新能源 汽车"],
        "游戏": ["游戏 新闻", "手游 主机游戏"],
        "育儿": ["育儿 新闻", "亲子 儿童 教育"],
        "职场": ["职场 新闻", "就业 招聘 职场"],
    },
)


GENERAL_FIXED_FEEDS: dict[str, list[str]] = {
    "国内": [
        "https://www.chinanews.com.cn/rss/scroll-news.xml",
        "https://www.chinanews.com.cn/rss/china.xml",
        "https://www.chinanews.com.cn/rss/society.xml",
    ],
    "国际": ["https://www.chinanews.com.cn/rss/world.xml"],
    "财经": ["https://www.chinanews.com.cn/rss/finance.xml"],
    "体育": ["https://www.chinanews.com.cn/rss/sports.xml"],
    "娱乐": ["https://www.chinanews.com.cn/rss/culture.xml"],
    "教育": ["https://www.chinanews.com.cn/rss/edu.xml"],
    "生活": ["https://www.chinanews.com.cn/rss/life.xml", "https://www.chinanews.com.cn/rss/jk.xml"],
}


def fetch_url(url: str, timeout: int = 10) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read()


def fetch_text(url: str, timeout: int = 10) -> str:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        data = response.read()
        content_type = response.headers.get("content-type", "")
    charset = "utf-8"
    match = re.search(r"charset=([\w.-]+)", content_type, re.I)
    if match:
        charset = match.group(1)
    try:
        return data.decode(charset, errors="replace")
    except LookupError:
        return data.decode("utf-8", errors="replace")


def text_from_html(value: str | None) -> str:
    if not value:
        return ""
    value = re.sub(r"(?is)<(script|style).*?</\1>", " ", value)
    value = re.sub(r"(?s)<[^>]+>", " ", value)
    value = html.unescape(value)
    return re.sub(r"\s+", " ", value).strip()


def extract_article_text(html_text: str) -> str:
    html_text = re.sub(
        r"(?is)<(script|style|noscript|svg|iframe|form|header|footer|nav|aside).*?</\1>",
        " ",
        html_text,
    )
    html_text = re.sub(r"(?is)<!--.*?-->", " ", html_text)

    candidates: list[str] = []
    for pattern in (
        r"(?is)<article\b[^>]*>(.*?)</article>",
        r"(?is)<main\b[^>]*>(.*?)</main>",
        r"(?is)<div\b[^>]+(?:class|id)=['\"][^'\"]*(?:article|content|post|story|entry)[^'\"]*['\"][^>]*>(.*?)</div>",
    ):
        candidates.extend(match.group(1) for match in re.finditer(pattern, html_text))
    if not candidates:
        candidates = [html_text]

    best = ""
    for candidate in candidates:
        parts = re.findall(r"(?is)<p\b[^>]*>(.*?)</p>|<h[12]\b[^>]*>(.*?)</h[12]>", candidate)
        paragraphs = []
        for part in parts:
            text = text_from_html(part[0] or part[1])
            if is_boilerplate_text(text):
                continue
            if len(text) >= 25:
                paragraphs.append(text)
        text = "\n".join(paragraphs)
        if len(text) > len(best):
            best = text

    if not best:
        best = text_from_html(html_text)
    best = re.sub(r"\n{3,}", "\n\n", best).strip()
    return best[:18000]


def is_boilerplate_text(text: str) -> bool:
    markers = (
        "即时 时政 财经",
        "关于我们 About us",
        "发表评论",
        "文明上网理性发言",
        "更多精彩内容请进入",
        "新闻精选",
        "编辑:",
        "【编辑:",
    )
    if any(marker in text for marker in markers):
        return True
    if text.count("|") >= 8:
        return True
    return False


def looks_chinese(value: str) -> bool:
    chinese_chars = len(re.findall(r"[\u4e00-\u9fff]", value))
    latin_chars = len(re.findall(r"[A-Za-z]", value))
    return chinese_chars > 0 and chinese_chars >= latin_chars * 0.2


def translate_text_to_chinese(text: str, errors: list[str]) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    return text


def split_sentences(text: str) -> list[str]:
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return []
    chunks = re.split(r"(?<=[。！？!?；;])\s*|(?<=[.])\s+(?=[A-Z0-9])", text)
    result = []
    for chunk in chunks:
        chunk = chunk.strip()
        if len(chunk) >= 18:
            result.append(chunk)
    return result


def summarize_to_chinese(
    config: ReportConfig,
    title: str,
    article_text: str,
    rss_summary: str,
    errors: list[str],
) -> tuple[str, list[str]]:
    source_text = rss_summary if len(rss_summary) >= 60 else article_text or rss_summary or title
    if is_boilerplate_text(source_text) and article_text and article_text != source_text:
        source_text = article_text
    sentences = split_sentences(source_text)
    if not sentences:
        sentences = [source_text]

    def sentence_score(sentence: str) -> float:
        lower = sentence.lower()
        keyword_score = sum(1 for keyword in config.keywords if keyword.lower() in lower) * 5
        length_score = min(len(sentence), 260) / 260
        return keyword_score + length_score

    ranked = sorted(sentences, key=sentence_score, reverse=True)
    selected = ranked[:8]
    translated = [translate_text_to_chinese(sentence, errors) for sentence in selected]

    points: list[str] = []
    for sentence in translated:
        sentence = re.sub(r"\s+", " ", sentence).strip()
        sentence = textwrap.shorten(sentence, width=180, placeholder="...")
        if sentence and sentence not in points:
            points.append(sentence)
        if len(points) >= 3:
            break

    summary = " ".join(points[:2])
    return summary, points


def parse_date(value: str | None) -> datetime | None:
    if not value:
        return None
    value = value.strip()
    try:
        parsed = parsedate_to_datetime(value)
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed
    except (TypeError, ValueError):
        pass
    for fmt in ("%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%d"):
        try:
            parsed = datetime.strptime(value, fmt)
            if parsed.tzinfo is None:
                parsed = parsed.replace(tzinfo=timezone.utc)
            return parsed
        except ValueError:
            continue
    return None


def localize(dt: datetime | None) -> str:
    if not dt:
        return "未知时间"
    return dt.astimezone().strftime("%Y-%m-%d %H:%M")


def clean_link(link: str) -> str:
    if "news.google.com" not in link:
        return link
    parsed = urllib.parse.urlparse(link)
    query = urllib.parse.parse_qs(parsed.query)
    for key in ("url", "u"):
        if query.get(key):
            return query[key][0]
    return link


def parse_rss(data: bytes, feed_url: str, category: str) -> list[NewsItem]:
    items: list[NewsItem] = []
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return items

    channel_title = root.findtext("./channel/title") or urllib.parse.urlparse(feed_url).netloc
    rss_items = root.findall(".//item")
    if not rss_items:
        rss_items = root.findall("{http://www.w3.org/2005/Atom}entry")

    for item in rss_items:
        if item.tag.endswith("entry"):
            title = item.findtext("{http://www.w3.org/2005/Atom}title") or ""
            link = ""
            for link_node in item.findall("{http://www.w3.org/2005/Atom}link"):
                link = link_node.attrib.get("href", "")
                if link:
                    break
            summary = (
                item.findtext("{http://www.w3.org/2005/Atom}summary")
                or item.findtext("{http://www.w3.org/2005/Atom}content")
                or ""
            )
            published = parse_date(
                item.findtext("{http://www.w3.org/2005/Atom}published")
                or item.findtext("{http://www.w3.org/2005/Atom}updated")
            )
            source = channel_title
        else:
            title = item.findtext("title") or ""
            link = item.findtext("link") or ""
            summary = item.findtext("description") or ""
            published = parse_date(item.findtext("pubDate") or item.findtext("date"))
            source = (
                item.findtext("source")
                or item.findtext("{http://search.yahoo.com/mrss/}credit")
                or channel_title
            )

        title = text_from_html(title)
        summary = text_from_html(summary)
        if title and link:
            items.append(
                NewsItem(
                    title=title,
                    link=clean_link(link.strip()),
                    source=text_from_html(source),
                    published=published,
                    summary=summary,
                    category=category,
                )
            )
    return items


def normalize_title(title: str) -> str:
    title = title.lower()
    title = re.sub(r"[^a-z0-9\u4e00-\u9fff]+", " ", title)
    return re.sub(r"\s+", " ", title).strip()


def build_search_feeds(query: str) -> list[str]:
    encoded = urllib.parse.quote(query)
    return [
        f"https://news.google.com/rss/search?q={encoded}%20when:2d&hl=zh-CN&gl=CN&ceid=CN:zh-Hans",
        f"https://www.bing.com/news/search?q={encoded}&format=rss",
    ]


def build_feeds(config: ReportConfig) -> list[tuple[str, str]]:
    feeds: list[tuple[str, str]] = []
    if config.key == "ai":
        for category, urls in AI_FIXED_FEEDS.items():
            feeds.extend((url, category) for url in urls)
        return feeds
    if config.key == "general":
        for category, urls in GENERAL_FIXED_FEEDS.items():
            feeds.extend((url, category) for url in urls)
        return feeds
    for term in config.search_terms:
        feeds.extend((url, "AI") for url in build_search_feeds(term))
    feeds.extend((url, "AI") for url in config.extra_feeds)
    return feeds


def score_item(config: ReportConfig, item: NewsItem) -> int:
    text = f"{item.title} {item.summary}".lower()
    score = 0
    for keyword in config.keywords:
        if keyword.lower() in text:
            score += 5
    if item.published:
        age_hours = max(
            0,
            (datetime.now(timezone.utc) - item.published.astimezone(timezone.utc)).total_seconds() / 3600,
        )
        score += max(0, 60 - int(age_hours))
    if item.summary:
        score += min(len(item.summary), 300) // 50
    return score


def collect_news(config: ReportConfig, limit: int) -> tuple[list[NewsItem], list[str]]:
    by_key: dict[str, NewsItem] = {}
    errors: list[str] = []
    for feed_url, category in build_feeds(config):
        try:
            for item in parse_rss(fetch_url(feed_url), feed_url, category):
                key = normalize_title(item.title)
                if not key:
                    continue
                existing = by_key.get(key)
                if existing is None or score_item(config, item) > score_item(config, existing):
                    by_key[key] = item
        except Exception as exc:
            errors.append(f"{feed_url}: {exc}")

    items = sorted(by_key.values(), key=lambda item: score_item(config, item), reverse=True)
    return items[:limit], errors


def enrich_news(config: ReportConfig, items: list[NewsItem], detail_limit: int, errors: list[str]) -> list[EnrichedNewsItem]:
    enriched: list[EnrichedNewsItem] = []
    for item in items[:detail_limit]:
        article_text = ""
        if len(item.summary) < 240:
            try:
                article_text = extract_article_text(fetch_text(item.link, timeout=8))
            except Exception as exc:
                errors.append(f"正文抓取失败：{item.title} - {exc}")

        title_cn = translate_text_to_chinese(item.title, errors)
        summary_cn, points = summarize_to_chinese(config, item.title, article_text, item.summary, errors)
        enriched.append(
            EnrichedNewsItem(
                item=item,
                article_text=article_text,
                chinese_title=title_cn,
                chinese_summary=summary_cn,
                key_points=points,
            )
        )
    return enriched


def xml_escape(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def paragraph(text: str, style: str | None = None) -> str:
    style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    text = xml_escape(text)
    lines = text.splitlines() or [""]
    runs = []
    for idx, line in enumerate(lines):
        if idx:
            runs.append("<w:r><w:br/></w:r>")
        runs.append(f'<w:r><w:t xml:space="preserve">{line}</w:t></w:r>')
    return f"<w:p>{style_xml}{''.join(runs)}</w:p>"


def limit_chars(text: str, max_chars: int = 900) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= max_chars:
        return text
    return text[: max_chars - 1].rstrip("，。；、 ") + "。"


def build_overall_summary(config: ReportConfig, items: list[EnrichedNewsItem]) -> str:
    if not items:
        return f"今日未抓取到可用的{config.title}内容，请检查网络连接或稍后重试。"

    category_counts: dict[str, int] = {}
    points: list[str] = []
    seen: set[str] = set()
    for enriched in items:
        category_counts[enriched.item.category] = category_counts.get(enriched.item.category, 0) + 1
        for point in enriched.key_points or [enriched.chinese_summary]:
            point = re.sub(r"\s+", " ", point).strip().rstrip("。！？!?")
            key = normalize_title(point)
            if point and key and key not in seen:
                seen.add(key)
                points.append(point)
            if len(points) >= 10:
                break
        if len(points) >= 10:
            break

    category_text = "、".join(f"{name}{count}条" for name, count in sorted(category_counts.items()))
    if points:
        return limit_chars(f"今日共整理{len(items)}条重点新闻，覆盖{category_text}。主要看点包括：" + "；".join(points[:8]) + "。")
    return f"今日共整理{len(items)}条重点新闻，覆盖{category_text}。部分来源未提供足够摘要，建议打开原文查看详细内容。"


def clean_report_summary(text: str) -> str:
    text = re.sub(r"https?://\S+", "", text)
    text = re.sub(r"\b[\w.-]+@[\w.-]+\.\w+\b", "", text)
    text = re.sub(r"\s+", " ", text).strip(" ：:，。；、-")
    return limit_chars(text, 220)


def dedupe_report_items(items: list[EnrichedNewsItem]) -> list[str]:
    summaries: list[str] = []
    seen: set[str] = set()
    for enriched in items:
        title = clean_report_summary(enriched.chinese_title)
        summary = clean_report_summary(enriched.chinese_summary or enriched.item.summary)
        if summary and title and normalize_title(title) not in normalize_title(summary):
            text = f"{title}：{summary}"
        else:
            text = summary or title
        key = normalize_title(text[:120])
        if not text or not key:
            continue
        if any(key in old or old in key for old in seen):
            continue
        seen.add(key)
        summaries.append(text)
    return summaries


def build_document_xml(config: ReportConfig, items: list[EnrichedNewsItem], errors: list[str], generated_at: datetime) -> str:
    body: list[str] = []
    body.append(paragraph(config.title, "Title"))
    summaries = dedupe_report_items(items)
    if summaries:
        for index, summary in enumerate(summaries, start=1):
            body.append(paragraph(f"{index}. {summary}"))
    else:
        body.append(paragraph(f"1. 今日未抓取到可用的{config.title}内容。"))

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {''.join(body)}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>"""


def write_docx(path: Path, config: ReportConfig, items: list[EnrichedNewsItem], errors: list[str], generated_at: datetime) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
    styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="22"/></w:rPr></w:style>
</w:styles>"""
    core = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>{xml_escape(config.title)}</dc:title>
  <dc:creator>AI News Automation</dc:creator>
  <dcterms:created>{generated_at.astimezone(timezone.utc).isoformat()}</dcterms:created>
</cp:coreProperties>"""
    app = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>AI News Automation</Application>
</Properties>"""

    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", build_document_xml(config, items, errors, generated_at))
        docx.writestr("word/styles.xml", styles)
        docx.writestr("docProps/core.xml", core)
        docx.writestr("docProps/app.xml", app)


def generate_report(config: ReportConfig, args: argparse.Namespace, generated_at: datetime, output_root: Path) -> Path:
    items, errors = collect_news(config, args.limit)
    enriched_items = enrich_news(config, items, args.detail_limit, errors)
    report_dir = output_root / generated_at.strftime("%Y-%m-%d")
    filename = f"{config.filename_prefix}_{generated_at.strftime('%Y%m%d_%H%M%S')}.docx"
    output_path = report_dir / filename
    write_docx(output_path, config, enriched_items, errors, generated_at)
    print(f"已生成：{output_path}")
    print(f"{config.title} 新闻数量：{len(enriched_items)}")
    if errors:
        print(f"{config.title} 抓取失败来源/正文数：{len(errors)}")
    return output_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate daily AI and general news Word reports.")
    parser.add_argument("--limit", type=int, default=40, help="Maximum number of feed items to collect per report.")
    parser.add_argument("--detail-limit", type=int, default=12, help="Maximum number of items to summarize per report.")
    parser.add_argument("--output-dir", default=str(OUTPUT_DIR), help="Root directory for generated .docx files.")
    parser.add_argument(
        "--report",
        choices=("all", "ai", "general"),
        default="all",
        help="Which report to generate.",
    )
    args = parser.parse_args()

    generated_at = datetime.now().astimezone()
    output_root = Path(args.output_dir)
    configs = []
    if args.report in ("all", "ai"):
        configs.append(AI_REPORT)
    if args.report in ("all", "general"):
        configs.append(GENERAL_REPORT)

    for config in configs:
        generate_report(config, args, generated_at, output_root)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
