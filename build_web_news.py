#!/usr/bin/env python3
"""
Pull public Google News RSS results and merge them into data.json news.

No API key, Gmail access, or Apps Script required.

Optional variables:
  WEB_NEWS_QUERIES       Pipe-separated queries. Defaults to Ultrahuman and
                         wearables/smart-ring sector queries.
  WEB_NEWS_LOOKBACK_DAYS Only keep items newer than this many days.
  WEB_NEWS_LIMIT         Max items to keep from all feeds combined.
"""
import datetime as dt
import email.utils
import html
import json
import os
import re
import sys
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path

DEFAULT_QUERIES = [
    "Ultrahuman",
    "wearables smart ring",
    "health wearable",
    "Oura OR WHOOP OR RingConn OR Ultrahuman",
]


def env(name):
    return os.environ.get(name, "").strip()


def int_env(name, default, lo=None, hi=None):
    try:
        value = int(env(name) or default)
    except ValueError:
        value = default
    if lo is not None:
        value = max(lo, value)
    if hi is not None:
        value = min(hi, value)
    return value


def queries():
    raw = env("WEB_NEWS_QUERIES")
    if raw:
        return [q.strip() for q in raw.split("|") if q.strip()]
    return DEFAULT_QUERIES


def feed_url(query):
    params = urllib.parse.urlencode({
        "q": query,
        "hl": "en-IN",
        "gl": "IN",
        "ceid": "IN:en",
    })
    return f"https://news.google.com/rss/search?{params}"


def fetch(url):
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0 UltrahumanDashboard/1.0"},
        method="GET",
    )
    with urllib.request.urlopen(req, timeout=45) as resp:
        return resp.read()


def clean(value):
    value = html.unescape(value or "")
    value = re.sub(r"<[^>]+>", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def parse_date(value):
    if not value:
        return None
    try:
        return email.utils.parsedate_to_datetime(value).date().isoformat()
    except Exception:
        return None


def item_text(item, tag):
    found = item.find(tag)
    return found.text if found is not None else ""


def item_source(item):
    source = item.find("source")
    return clean(source.text if source is not None else "") or "Google News"


def is_recent(date_text):
    days = int_env("WEB_NEWS_LOOKBACK_DAYS", 30, lo=1, hi=365)
    parsed = parse_date(date_text)
    if not parsed:
        return True
    date = dt.date.fromisoformat(parsed)
    return date >= dt.date.today() - dt.timedelta(days=days)


def classify(title, query):
    text = f"{title} {query}".lower()
    if "ultrahuman" in text:
        return "Company", "Ultrahuman"
    return "Industry", ""


def read_feed(query):
    xml_bytes = fetch(feed_url(query))
    root = ET.fromstring(xml_bytes)
    out = []
    for item in root.findall("./channel/item"):
        pub = item_text(item, "pubDate")
        if not is_recent(pub):
            continue
        title = clean(item_text(item, "title"))
        link = clean(item_text(item, "link"))
        desc = clean(item_text(item, "description"))
        if not title or not link:
            continue
        typ, company = classify(title, query)
        out.append({
            "date": parse_date(pub),
            "type": typ,
            "company": company,
            "theme": "Wearables",
            "source": item_source(item),
            "headline": title,
            "url": link,
            "summary": desc if desc != title else "",
            "sentiment": "",
            "relevance": "Web News",
        })
    return out


def item_key(item):
    url = item.get("url") or ""
    title = re.sub(r"\W+", "", (item.get("headline") or "").lower())
    return url or title


def merge_news(data, web_items):
    news = data.get("news") or {"items": [], "industry": [], "companies": []}
    existing = news.get("items") or []
    seen = {item_key(x) for x in existing if item_key(x)}
    merged = list(existing)
    added = 0
    for item in web_items:
        key = item_key(item)
        if not key or key in seen:
            continue
        seen.add(key)
        merged.append(item)
        added += 1
    merged.sort(key=lambda x: x.get("date") or "", reverse=True)
    limit = int_env("WEB_NEWS_LIMIT", 30, lo=1, hi=100)
    merged = merged[:limit]
    news["items"] = merged
    news["industry"] = [x for x in merged if str(x.get("type", "")).lower() == "industry" or not x.get("company")]
    news["companies"] = [x for x in merged if str(x.get("type", "")).lower() == "company" or x.get("company")]
    news["meta"] = {
        "webNewsQueries": queries(),
        "webNewsItemsSeen": len(web_items),
        "webNewsItemsAdded": added,
    }
    data["news"] = news


def main():
    out_path = sys.argv[1] if len(sys.argv) > 1 else "data.json"
    data = json.loads(Path(out_path).read_text()) if Path(out_path).exists() else {}
    items = []
    for query in queries():
        try:
            items.extend(read_feed(query))
        except Exception as e:
            print(f"! skipped web news query {query!r}: {e}")
    merge_news(data, items)
    Path(out_path).write_text(json.dumps(data, indent=2, ensure_ascii=False))
    print(f"✓ merged {len(items)} web news items into news from {len(queries())} queries")


if __name__ == "__main__":
    main()
