#!/usr/bin/env python3
"""
build_meetings.py — pull recent Cashify meeting notes from Notion into data.json.

Requires NOTION_TOKEN or NOTION_API_KEY. If no token is present, preserves any
existing meetings block and exits without failing the dashboard build.
"""
import datetime
import json
import os
import re
import sys
import urllib.error
import urllib.request
from pathlib import Path

NOTION_VERSION = "2022-06-28"
SEARCH_URL = "https://api.notion.com/v1/search"
MEETING_TERMS = re.compile(r"\b(meeting|discussion|call|banker|roadmap|aop|board)\b", re.I)
DATE_PATTERNS = [
    re.compile(r"(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})"),
    re.compile(r"(\d{1,2})[- ]([A-Za-z]{3,9})[- ](\d{2,4})"),
]
MONTHS = {
    "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
    "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9, "oct": 10,
    "october": 10, "nov": 11, "november": 11, "dec": 12, "december": 12,
}


def notion_token():
    return os.environ.get("NOTION_TOKEN") or os.environ.get("NOTION_API_KEY")


def request_json(url, token, body=None):
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(
        url,
        data=data,
        headers={
            "Authorization": f"Bearer {token}",
            "Notion-Version": NOTION_VERSION,
            "Content-Type": "application/json",
        },
        method="POST" if body is not None else "GET",
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read().decode("utf-8"))


def rich_text(parts):
    return "".join(p.get("plain_text", "") for p in parts or []).strip()


def page_title(page):
    props = page.get("properties", {})
    for prop in props.values():
        if prop.get("type") == "title":
            return rich_text(prop.get("title", []))
    return ""


def page_url(page):
    return page.get("url", "")


def parse_title_date(title):
    for pat in DATE_PATTERNS:
        m = pat.search(title)
        if not m:
            continue
        d, mid, y = m.groups()
        if mid.isdigit():
            month = int(mid)
        else:
            month = MONTHS.get(mid.lower())
            if not month:
                continue
        year = int(y)
        if year < 100:
            year += 2000
        try:
            return datetime.date(year, month, int(d)).isoformat()
        except ValueError:
            continue
    return None


def block_text(block):
    typ = block.get("type")
    obj = block.get(typ, {}) if typ else {}
    return rich_text(obj.get("rich_text", []))


def fetch_block_children(token, block_id, max_blocks=80):
    blocks, cursor = [], None
    while len(blocks) < max_blocks:
        url = f"https://api.notion.com/v1/blocks/{block_id}/children?page_size=100"
        if cursor:
            url += f"&start_cursor={cursor}"
        data = request_json(url, token)
        blocks.extend(data.get("results", []))
        if not data.get("has_more"):
            break
        cursor = data.get("next_cursor")
    return blocks[:max_blocks]


def extract_summary(token, page_id):
    bullets, sections = [], []
    for block in fetch_block_children(token, page_id):
        typ = block.get("type")
        text = block_text(block)
        if not text:
            continue
        if typ in ("heading_1", "heading_2", "heading_3"):
            sections.append(text.strip("# *"))
        elif typ in ("bulleted_list_item", "numbered_list_item", "paragraph", "to_do"):
            clean = text.strip("•- \t")
            if clean and len(clean) > 8:
                bullets.append(clean)
        if len(bullets) >= 6:
            break
    return {"sections": sections[:4], "bullets": bullets[:5]}


def search_meetings(token):
    query = os.environ.get("NOTION_MEETING_QUERY", "Cashify")
    wanted = int(os.environ.get("NOTION_MEETING_LIMIT", "5"))
    body = {
        "query": query,
        "filter": {"property": "object", "value": "page"},
        "sort": {"direction": "descending", "timestamp": "last_edited_time"},
        "page_size": min(25, max(10, wanted * 3)),
    }
    data = request_json(SEARCH_URL, token, body)
    out = []
    for page in data.get("results", []):
        title = page_title(page)
        if "cashify" not in title.lower() or not MEETING_TERMS.search(title):
            continue
        summary = extract_summary(token, page["id"])
        out.append({
            "title": title,
            "date": parse_title_date(title),
            "lastEdited": page.get("last_edited_time"),
            "url": page_url(page),
            "sections": summary["sections"],
            "bullets": summary["bullets"],
        })
        if len(out) >= wanted:
            break
    out.sort(key=lambda x: x.get("date") or x.get("lastEdited") or "", reverse=True)
    return out


def main():
    out_path = sys.argv[1] if len(sys.argv) > 1 else "data.json"
    data = json.loads(Path(out_path).read_text()) if Path(out_path).exists() else {}
    token = notion_token()
    if not token:
        data.setdefault("meetings", {"items": []})
        Path(out_path).write_text(json.dumps(data, indent=2, ensure_ascii=False))
        print("! skipped Notion meetings: set NOTION_TOKEN or NOTION_API_KEY")
        return
    try:
        items = search_meetings(token)
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        sys.exit(f"ERROR: Notion API failed HTTP {e.code}: {body}")
    data["meetings"] = {"items": items}
    Path(out_path).write_text(json.dumps(data, indent=2, ensure_ascii=False))
    print(f"✓ wrote {len(items)} recent Notion meetings to {out_path}")


if __name__ == "__main__":
    main()
