#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Daily news collector.

Generates two Word reports:
- AI相关新闻日报
- 大杂烩新闻日报

Reports are saved under <script_dir>/output/yyyy-MM-dd by default.
Each report also gets a fixed "latest" copy under <script_dir>/output/最新/.

Optional features (configured via .env or environment variables):
- LLM-powered Chinese summaries (AI_NEWS_API_KEY, AI_NEWS_API_BASE, AI_NEWS_MODEL);
  falls back to local extractive summaries when not configured.
- Push notifications via Server酱 / 企业微信机器人 / SMTP email
  (AI_NEWS_SERVERCHAN_KEY, AI_NEWS_WECOM_WEBHOOK, AI_NEWS_SMTP_*).
"""

from __future__ import annotations

import argparse
import html
import json
import os
import re
import shutil
import smtplib
import ssl
import textwrap
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from difflib import SequenceMatcher
from email.message import EmailMessage
from email.utils import parsedate_to_datetime
from html.parser import HTMLParser
from pathlib import Path
from xml.etree import ElementTree as ET


OUTPUT_DIR = Path(__file__).resolve().parent / "output"
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
    rss_text: str
    category: str


@dataclass(frozen=True)
class EnrichedNewsItem:
    item: NewsItem
    article_text: str
    chinese_title: str
    chinese_summary: str
    key_points: list[str]


@dataclass(frozen=True)
class LLMConfig:
    api_key: str = ""
    api_base: str = "https://api.deepseek.com"
    model: str = "deepseek-chat"

    @property
    def enabled(self) -> bool:
        return bool(self.api_key)


def env_get(env: dict[str, str], name: str) -> str:
    """Read a value from the loaded .env dict, falling back to real env vars."""
    return env.get(name) or os.environ.get(name, "")


def load_env(path: Path | None = None) -> dict[str, str]:
    """Load KEY=VALUE pairs from a .env file (stdlib only)."""
    env: dict[str, str] = {}
    candidates = [path, Path(__file__).resolve().parent / ".env"]
    for candidate in candidates:
        if not candidate or not candidate.is_file():
            continue
        for raw in candidate.read_text(encoding="utf-8", errors="ignore").splitlines():
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            env[key.strip()] = value.strip().strip('"').strip("'")
    return env


def call_llm(llm: LLMConfig, messages: list[dict], errors: list[str], timeout: int = 60) -> str | None:
    """Call an OpenAI-compatible chat completions API; returns the assistant text or None."""
    if not llm.enabled:
        return None
    payload = json.dumps(
        {"model": llm.model, "messages": messages, "temperature": 0.3, "max_tokens": 800}
    ).encode("utf-8")
    base = llm.api_base.rstrip("/")
    url = base if base.endswith("/chat/completions") else f"{base}/chat/completions"
    request = urllib.request.Request(
        url,
        data=payload,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {llm.api_key}",
            "User-Agent": USER_AGENT,
        },
    )
    for attempt in range(2):
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                data = json.loads(response.read().decode("utf-8"))
            return data["choices"][0]["message"]["content"]
        except Exception as exc:
            if attempt == 1:
                errors.append(f"LLM 调用失败：{exc}")
    return None


def parse_llm_json(text: str) -> dict | None:
    """Parse a JSON object from an LLM reply, tolerating markdown fences."""
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    try:
        data = json.loads(text)
        return data if isinstance(data, dict) else None
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", text, re.S)
        if match:
            try:
                data = json.loads(match.group(0))
                return data if isinstance(data, dict) else None
            except json.JSONDecodeError:
                return None
    return None


def llm_enrich(
    llm: LLMConfig,
    item: NewsItem,
    article_text: str,
    errors: list[str],
) -> tuple[str, str, list[str]] | None:
    """One-shot LLM call producing a Chinese title, summary and key points."""
    if not llm.enabled:
        return None
    source = clean_text(article_text) or item.summary or item.title
    prompt = (
        "你是一名资深中文新闻编辑。请根据下面的新闻标题和原文完成三项任务，"
        "并只输出一个 JSON 对象（不要输出任何其他文字）：\n"
        '{"chinese_title": "中文标题，英文标题翻译为中文并润色，保留专有名词原名", '
        '"summary": "约120字的中文摘要，客观概括新闻要点", '
        '"key_points": ["要点1", "要点2", "要点3"]}\n\n'
        f"新闻标题：{item.title}\n\n新闻原文：\n{source[:3000]}"
    )
    content = call_llm(
        llm,
        [
            {"role": "system", "content": "你只输出 JSON，不要输出其他内容。"},
            {"role": "user", "content": prompt},
        ],
        errors,
    )
    if not content:
        return None
    data = parse_llm_json(content)
    if not data:
        return None
    title = clean_report_summary(str(data.get("chinese_title") or "")) or item.title
    summary = clean_report_summary(str(data.get("summary") or ""))
    if not summary:
        return None
    points_raw = data.get("key_points") or []
    points = [clean_report_summary(str(point)) for point in points_raw if str(point).strip()]
    if not points:
        points = [summary]
    return title, summary, points[:5]


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

AI_STRONG_KEYWORDS = (
    "ai",
    "人工智能",
    "大模型",
    "生成式",
    "模型",
    "openai",
    "chatgpt",
    "gpt",
    "claude",
    "anthropic",
    "deepseek",
    "gemini",
    "智能体",
    "agent",
    "算力",
    "芯片",
    "nvidia",
    "英伟达",
    "机器人",
    "自动驾驶",
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


def repair_mojibake(value: str) -> str:
    """Repair common UTF-8 text that was decoded as Latin-1/GBK."""
    if not value:
        return ""
    candidates = [value]
    for wrong_encoding in ("latin1", "cp1252", "gbk"):
        try:
            candidates.append(value.encode(wrong_encoding, errors="ignore").decode("utf-8", errors="ignore"))
        except (LookupError, UnicodeError):
            pass

    def badness(text: str) -> int:
        return (
            text.count("�") * 8
            + len(re.findall(r"[锟斤拷]", text)) * 6
            + len(re.findall(r"[æäåçèéã]", text, flags=re.I)) * 2
        )

    return min(candidates, key=badness).strip()


def text_from_html(value: str | None) -> str:
    if not value:
        return ""
    value = re.sub(r"(?is)<(script|style).*?</\1>", " ", value)
    value = re.sub(r"(?s)<[^>]+>", " ", value)
    value = html.unescape(value)
    return clean_text(value)


def clean_text(value: str | None) -> str:
    if not value:
        return ""
    value = html.unescape(value)
    value = repair_mojibake(value)
    value = value.replace("\u3000", " ")
    value = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", " ", value)
    value = re.sub(r"[ \t\r\f\v]+", " ", value)
    value = re.sub(r"\n[ \t]+", "\n", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


class ReadabilityParser(HTMLParser):
    BLOCK_TAGS = {"p", "div", "article", "section", "main", "li", "blockquote", "h1", "h2", "h3"}
    SKIP_TAGS = {"script", "style", "noscript", "svg", "iframe", "form", "header", "footer", "nav", "aside"}
    SKIP_HINTS = (
        "nav",
        "menu",
        "footer",
        "header",
        "comment",
        "related",
        "share",
        "social",
        "advert",
        "subscribe",
        "breadcrumb",
        "copyright",
    )
    GOOD_HINTS = ("article", "content", "post", "story", "entry", "main", "news", "正文", "内容")

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.blocks: list[tuple[str, int]] = []
        self.current: list[str] = []
        self.depth = 0
        self.skip_depth = 0
        self.good_depth = 0
        self.tag_stack: list[tuple[str, bool, bool]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr_text = " ".join(value or "" for name, value in attrs if name in {"class", "id", "role"}).lower()
        skip_this = tag in self.SKIP_TAGS or any(hint in attr_text for hint in self.SKIP_HINTS)
        good_this = any(hint in attr_text for hint in self.GOOD_HINTS)
        self.tag_stack.append((tag, skip_this, good_this))
        if skip_this:
            self.skip_depth += 1
        if good_this:
            self.good_depth += 1
        if tag in self.BLOCK_TAGS:
            self.flush()
            self.depth += 1
        if tag == "br":
            self.current.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in self.BLOCK_TAGS:
            self.flush()
            self.depth = max(0, self.depth - 1)
        while self.tag_stack:
            open_tag, skip_this, good_this = self.tag_stack.pop()
            if skip_this and self.skip_depth:
                self.skip_depth -= 1
            if good_this and self.good_depth:
                self.good_depth -= 1
            if open_tag == tag:
                break

    def handle_data(self, data: str) -> None:
        if not self.skip_depth:
            self.current.append(data)

    def flush(self) -> None:
        text = clean_text("".join(self.current))
        self.current = []
        if len(text) >= 20 and not is_boilerplate_text(text):
            bonus = 120 if self.good_depth else 0
            self.blocks.append((text, bonus))

    def close(self) -> None:
        super().close()
        self.flush()


def extract_article_text(html_text: str) -> str:
    parser = ReadabilityParser()
    try:
        parser.feed(html_text)
        parser.close()
    except Exception:
        return text_from_html(html_text)[:18000]

    ranked: list[tuple[int, int, str]] = []
    seen: set[str] = set()
    for index, (text, bonus) in enumerate(parser.blocks):
        text = clean_text(text)
        key = normalize_title(text[:160])
        if not key or key in seen or is_boilerplate_text(text):
            continue
        seen.add(key)
        punctuation = text.count("。") + text.count("，") + text.count(".") + text.count(",")
        chinese_chars = len(re.findall(r"[\u4e00-\u9fff]", text))
        alpha_chars = len(re.findall(r"[A-Za-z]", text))
        score = len(text) + punctuation * 12 + bonus
        if chinese_chars or alpha_chars:
            score += min(chinese_chars + alpha_chars, 500)
        ranked.append((score, index, text))

    if not ranked:
        return text_from_html(html_text)[:18000]

    useful = [
        (index, text)
        for score, index, text in ranked
        if score >= 80 and len(text) >= 30
    ]
    useful.sort(key=lambda item: item[0])
    article = "\n".join(text for _, text in useful)
    article = clean_text(article)
    return article[:18000]


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
        "责任编辑",
        "免责声明",
        "copyright",
        "all rights reserved",
        "click here",
        "read more",
        "subscribe",
        "sign up",
        "广告",
        "相关阅读",
    )
    lower = text.lower()
    if any(marker.lower() in lower for marker in markers):
        return True
    if text.count("|") >= 8:
        return True
    if len(text) < 18:
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
    article_text = clean_text(article_text)
    rss_summary = clean_text(rss_summary)
    source_text = article_text if len(article_text) >= max(220, len(rss_summary) * 2) else rss_summary or article_text or title
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


def find_text_any(node: ET.Element, paths: tuple[str, ...]) -> str:
    for path in paths:
        value = node.findtext(path)
        if value:
            return value
    return ""


def find_atom_link(item: ET.Element) -> str:
    for link_node in item.findall("{http://www.w3.org/2005/Atom}link"):
        rel = link_node.attrib.get("rel", "alternate")
        href = link_node.attrib.get("href", "")
        if href and rel == "alternate":
            return href
    for link_node in item.findall("{http://www.w3.org/2005/Atom}link"):
        href = link_node.attrib.get("href", "")
        if href:
            return href
    return ""


def extract_rss_body(item: ET.Element, is_atom: bool) -> tuple[str, str]:
    if is_atom:
        summary_raw = find_text_any(
            item,
            (
                "{http://www.w3.org/2005/Atom}summary",
                "{http://www.w3.org/2005/Atom}subtitle",
            ),
        )
        content_raw = find_text_any(item, ("{http://www.w3.org/2005/Atom}content",))
    else:
        summary_raw = find_text_any(
            item,
            (
                "description",
                "{http://purl.org/rss/1.0/modules/content/}description",
                "{http://search.yahoo.com/mrss/}description",
            ),
        )
        content_raw = find_text_any(
            item,
            (
                "{http://purl.org/rss/1.0/modules/content/}encoded",
                "content",
                "{http://www.w3.org/2005/Atom}content",
            ),
        )
    summary = text_from_html(summary_raw)
    content = text_from_html(content_raw)
    if len(content) < len(summary):
        content = summary
    return summary, content


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
        is_atom = item.tag.endswith("entry")
        if is_atom:
            title = item.findtext("{http://www.w3.org/2005/Atom}title") or ""
            link = find_atom_link(item)
            published = parse_date(
                item.findtext("{http://www.w3.org/2005/Atom}published")
                or item.findtext("{http://www.w3.org/2005/Atom}updated")
            )
            source = channel_title
        else:
            title = item.findtext("title") or ""
            link = item.findtext("link") or ""
            published = parse_date(item.findtext("pubDate") or item.findtext("date"))
            source = (
                item.findtext("source")
                or item.findtext("{http://search.yahoo.com/mrss/}credit")
                or channel_title
            )

        title = text_from_html(title)
        summary, rss_text = extract_rss_body(item, is_atom)
        link = clean_link(html.unescape(link.strip()))
        if title and link:
            items.append(
                NewsItem(
                    title=title,
                    link=link,
                    source=text_from_html(source),
                    published=published,
                    summary=summary,
                    rss_text=rss_text,
                    category=category,
                )
            )
    return items


def normalize_title(title: str) -> str:
    title = repair_mojibake(title).lower()
    title = re.sub(r"[^a-z0-9\u4e00-\u9fff]+", " ", title)
    return re.sub(r"\s+", " ", title).strip()


def canonical_link(link: str) -> str:
    parsed = urllib.parse.urlparse(link)
    query = urllib.parse.parse_qsl(parsed.query, keep_blank_values=False)
    filtered = [
        (key, value)
        for key, value in query
        if not key.lower().startswith("utm_")
        and key.lower() not in {"spm", "from", "source", "ref", "ref_src", "ocid"}
    ]
    return urllib.parse.urlunparse(
        (
            parsed.scheme.lower(),
            parsed.netloc.lower().removeprefix("www."),
            parsed.path.rstrip("/"),
            "",
            urllib.parse.urlencode(filtered),
            "",
        )
    )


def title_tokens(title: str) -> set[str]:
    normalized = normalize_title(title)
    tokens = set(re.findall(r"[a-z0-9]{3,}|[\u4e00-\u9fff]{2,}", normalized))
    return tokens


def similar_titles(left: str, right: str) -> bool:
    left_norm = normalize_title(left)
    right_norm = normalize_title(right)
    if not left_norm or not right_norm:
        return False
    if left_norm in right_norm or right_norm in left_norm:
        return True
    ratio = SequenceMatcher(None, left_norm, right_norm).ratio()
    if ratio >= 0.82:
        return True
    left_tokens = title_tokens(left)
    right_tokens = title_tokens(right)
    if not left_tokens or not right_tokens:
        return False
    overlap = len(left_tokens & right_tokens) / max(1, min(len(left_tokens), len(right_tokens)))
    return overlap >= 0.75


def content_length(item: NewsItem) -> int:
    return max(len(item.rss_text), len(item.summary))


def choose_better_item(config: ReportConfig, left: NewsItem, right: NewsItem) -> NewsItem:
    left_score = score_item(config, left) + min(content_length(left), 3000) // 80
    right_score = score_item(config, right) + min(content_length(right), 3000) // 80
    if left_score != right_score:
        return left if left_score > right_score else right
    left_time = left.published.timestamp() if left.published else 0
    right_time = right.published.timestamp() if right.published else 0
    return left if left_time >= right_time else right


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
    text = f"{item.title} {item.summary} {item.rss_text[:1000]}".lower()
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
    if item.rss_text:
        score += min(len(item.rss_text), 1200) // 120
    return score


def is_relevant_item(config: ReportConfig, item: NewsItem) -> bool:
    if config.key != "ai":
        return True
    text = f"{item.title} {item.summary} {item.rss_text[:1200]}".lower()
    return any(keyword.lower() in text for keyword in AI_STRONG_KEYWORDS)


def collect_news(config: ReportConfig, limit: int) -> tuple[list[NewsItem], list[str]]:
    items: list[NewsItem] = []
    errors: list[str] = []
    for feed_url, category in build_feeds(config):
        try:
            for item in parse_rss(fetch_url(feed_url), feed_url, category):
                if not normalize_title(item.title):
                    continue
                if not is_relevant_item(config, item):
                    continue
                replacement_index: int | None = None
                item_link = canonical_link(item.link)
                for index, existing in enumerate(items):
                    same_link = canonical_link(existing.link) == item_link
                    same_title = similar_titles(existing.title, item.title)
                    if same_link or same_title:
                        replacement_index = index
                        break
                if replacement_index is None:
                    items.append(item)
                else:
                    items[replacement_index] = choose_better_item(config, items[replacement_index], item)
        except Exception as exc:
            errors.append(f"{feed_url}: {exc}")

    items = sorted(items, key=lambda item: score_item(config, item), reverse=True)
    return items[:limit], errors


def enrich_news(
    config: ReportConfig,
    items: list[NewsItem],
    detail_limit: int,
    errors: list[str],
    llm: LLMConfig | None = None,
) -> list[EnrichedNewsItem]:
    enriched: list[EnrichedNewsItem] = []
    for item in items[:detail_limit]:
        article_text = item.rss_text
        try:
            fetched_text = extract_article_text(fetch_text(item.link, timeout=10))
            if len(fetched_text) > len(article_text):
                article_text = fetched_text
        except Exception as exc:
            errors.append(f"正文抓取失败：{item.title} - {exc}")

        llm_result = llm_enrich(llm, item, article_text, errors) if llm and llm.enabled else None
        if llm_result is not None:
            title_cn, summary_cn, points = llm_result
        else:
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


def dedupe_report_items(items: list[EnrichedNewsItem]) -> list[EnrichedNewsItem]:
    summaries: list[str] = []
    result: list[EnrichedNewsItem] = []
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
        if any(key in old or old in key or similar_titles(text, old) for old in seen):
            continue
        seen.add(key)
        summaries.append(text)
        result.append(enriched)
    return result


def build_document_xml(config: ReportConfig, items: list[EnrichedNewsItem], errors: list[str], generated_at: datetime) -> str:
    body: list[str] = []
    body.append(paragraph(config.title, "Title"))
    body.append(paragraph(f"生成时间：{generated_at.strftime('%Y-%m-%d %H:%M:%S')}"))
    body.append(paragraph(build_overall_summary(config, items)))
    report_items = dedupe_report_items(items)
    if report_items:
        for index, enriched in enumerate(report_items, start=1):
            title = clean_report_summary(enriched.chinese_title)
            summary = clean_report_summary(enriched.chinese_summary or enriched.item.summary)
            if summary and title and normalize_title(title) not in normalize_title(summary):
                summary = f"{title}：{summary}"
            else:
                summary = summary or title
            body.append(paragraph(f"{index}. {summary}"))
            body.append(paragraph(f"来源：{enriched.item.source} | 分类：{enriched.item.category} | 时间：{localize(enriched.item.published)}"))
            body.append(paragraph(f"链接：{enriched.item.link}"))
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


def generate_report(
    config: ReportConfig,
    args: argparse.Namespace,
    generated_at: datetime,
    output_root: Path,
    llm: LLMConfig,
    env: dict[str, str],
) -> Path:
    items, errors = collect_news(config, args.limit)
    enriched_items = enrich_news(config, items, args.detail_limit, errors, llm)
    report_dir = output_root / generated_at.strftime("%Y-%m-%d")
    filename = f"{config.filename_prefix}_{generated_at.strftime('%Y-%m-%d')}.docx"
    output_path = report_dir / filename
    write_docx(output_path, config, enriched_items, errors, generated_at)
    print(f"已生成：{output_path}")
    print(f"{config.title} 新闻数量：{len(enriched_items)}")
    if errors:
        print(f"{config.title} 抓取失败来源/正文数：{len(errors)}")

    latest_dir = output_root / "最新"
    latest_path = latest_dir / f"{config.filename_prefix}_最新.docx"
    latest_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(output_path, latest_path)
    print(f"最新版：{latest_path}")

    if not args.no_push:
        push_report(config, enriched_items, output_path, generated_at, env, errors)
    return output_path


def build_push_digest(config: ReportConfig, items: list[EnrichedNewsItem], generated_at: datetime) -> str:
    lines = [
        f"# {config.title}",
        f"生成时间：{generated_at.strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        build_overall_summary(config, items),
        "",
    ]
    for index, enriched in enumerate(dedupe_report_items(items)[:8], start=1):
        title = clean_report_summary(enriched.chinese_title)
        summary = clean_report_summary(enriched.chinese_summary or enriched.item.summary)
        lines.append(f"{index}. {title}")
        if summary and title and normalize_title(title) not in normalize_title(summary):
            lines.append(f"　{summary}")
        lines.append(f"　来源：{enriched.item.source}｜链接：{enriched.item.link}")
        lines.append("")
    return "\n".join(lines)


def push_via_serverchan(env: dict[str, str], title: str, digest: str, errors: list[str]) -> bool:
    key = env_get(env, "AI_NEWS_SERVERCHAN_KEY")
    if not key:
        return False
    data = urllib.parse.urlencode({"title": title, "desp": digest}).encode("utf-8")
    request = urllib.request.Request(
        f"https://sctapi.ftqq.com/{key}.send",
        data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded", "User-Agent": USER_AGENT},
    )
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            result = json.loads(response.read().decode("utf-8"))
        if result.get("code") == 0:
            print("已推送：Server酱")
            return True
        errors.append(f"Server酱推送失败：{result}")
    except Exception as exc:
        errors.append(f"Server酱推送失败：{exc}")
    return False


def push_via_wecom(env: dict[str, str], title: str, digest: str, errors: list[str]) -> bool:
    webhook = env_get(env, "AI_NEWS_WECOM_WEBHOOK")
    if not webhook:
        return False
    payload = json.dumps(
        {"msgtype": "markdown", "markdown": {"content": f"{title}\n\n{digest}"[:4000]}}
    ).encode("utf-8")
    request = urllib.request.Request(
        webhook,
        data=payload,
        headers={"Content-Type": "application/json", "User-Agent": USER_AGENT},
    )
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            result = json.loads(response.read().decode("utf-8"))
        if result.get("errcode") == 0:
            print("已推送：企业微信机器人")
            return True
        errors.append(f"企业微信推送失败：{result}")
    except Exception as exc:
        errors.append(f"企业微信推送失败：{exc}")
    return False


def push_via_email(env: dict[str, str], title: str, digest: str, docx_path: Path, errors: list[str]) -> bool:
    host = env_get(env, "AI_NEWS_SMTP_HOST")
    user = env_get(env, "AI_NEWS_SMTP_USER")
    password = env_get(env, "AI_NEWS_SMTP_PASS")
    to_addrs = env_get(env, "AI_NEWS_SMTP_TO")
    if not (host and user and password and to_addrs):
        return False
    port = int(env_get(env, "AI_NEWS_SMTP_PORT") or "465")
    message = EmailMessage()
    message["Subject"] = title
    message["From"] = user
    message["To"] = to_addrs
    message.set_content(digest)
    if docx_path and docx_path.exists():
        message.add_attachment(
            docx_path.read_bytes(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=docx_path.name,
        )
    try:
        context = ssl.create_default_context()
        if port == 587:
            with smtplib.SMTP(host, port, timeout=30) as server:
                server.starttls(context=context)
                server.login(user, password)
                server.send_message(message)
        else:
            with smtplib.SMTP_SSL(host, port, timeout=30, context=context) as server:
                server.login(user, password)
                server.send_message(message)
        print("已推送：邮件")
        return True
    except Exception as exc:
        errors.append(f"邮件推送失败：{exc}")
    return False


def push_report(
    config: ReportConfig,
    items: list[EnrichedNewsItem],
    output_path: Path,
    generated_at: datetime,
    env: dict[str, str],
    errors: list[str],
) -> None:
    """Push the digest through every configured channel; failures are logged, never fatal."""
    channels = [
        ("Server酱", bool(env_get(env, "AI_NEWS_SERVERCHAN_KEY")), push_via_serverchan),
        ("企业微信", bool(env_get(env, "AI_NEWS_WECOM_WEBHOOK")), push_via_wecom),
        ("邮件", bool(env_get(env, "AI_NEWS_SMTP_HOST")), push_via_email),
    ]
    configured = [name for name, is_set, _ in channels if is_set]
    if not configured:
        print("未配置推送渠道（.env 中设置 AI_NEWS_SERVERCHAN_KEY / AI_NEWS_WECOM_WEBHOOK / AI_NEWS_SMTP_* 可启用），跳过推送")
        return
    title = f"{config.title} {generated_at.strftime('%Y-%m-%d')}"
    digest = build_push_digest(config, items, generated_at)
    for name, is_set, channel in channels:
        if not is_set:
            continue
        try:
            if name == "邮件":
                channel(env, title, digest, output_path, errors)
            else:
                channel(env, title, digest, errors)
        except Exception as exc:
            errors.append(f"{name} 推送异常：{exc}")


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
    parser.add_argument("--api-key", help="LLM API Key（默认读取 .env 的 AI_NEWS_API_KEY）")
    parser.add_argument("--api-base", help="LLM API 地址（默认 https://api.deepseek.com）")
    parser.add_argument("--model", help="LLM 模型名（默认 deepseek-chat）")
    parser.add_argument("--no-llm", action="store_true", help="禁用 LLM 摘要，使用本地抽取式摘要")
    parser.add_argument("--no-push", action="store_true", help="禁用日报推送")
    args = parser.parse_args()

    env = load_env()
    api_key = args.api_key or env_get(env, "AI_NEWS_API_KEY")
    llm = LLMConfig(
        api_key=api_key,
        api_base=args.api_base or env_get(env, "AI_NEWS_API_BASE") or "https://api.deepseek.com",
        model=args.model or env_get(env, "AI_NEWS_MODEL") or "deepseek-chat",
    )
    if not llm.enabled:
        print("提示：未配置 LLM API Key（.env 中 AI_NEWS_API_KEY），将使用本地抽取式摘要")
    elif args.no_llm:
        print("已通过 --no-llm 禁用 LLM，使用本地抽取式摘要")
        llm = LLMConfig()

    generated_at = datetime.now().astimezone()
    output_root = Path(args.output_dir)
    configs = []
    if args.report in ("all", "ai"):
        configs.append(AI_REPORT)
    if args.report in ("all", "general"):
        configs.append(GENERAL_REPORT)

    for config in configs:
        generate_report(config, args, generated_at, output_root, llm, env)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
