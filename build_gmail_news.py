#!/usr/bin/env python3
"""
Pull Google Alerts from Gmail and merge them into data.json news.

Required GitHub secrets:
  GMAIL_CLIENT_ID
  GMAIL_CLIENT_SECRET
  GMAIL_REFRESH_TOKEN

Optional variables:
  GMAIL_NEWS_QUERY       Gmail search query. Defaults to Google Alerts with
                         Ultrahuman OR Wearables over the configured lookback.
  GMAIL_NEWS_LOOKBACK_DAYS
  GMAIL_NEWS_LIMIT
"""
import base64
import datetime as dt
import email.utils
import html
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

TOKEN_URL = "https://oauth2.googleapis.com/token"
GMAIL_API = "https://gmail.googleapis.com/gmail/v1/users/me"
GOOGLE_ALERTS_FROM = "googlealerts-noreply@google.com"


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


def request_json(url, method="GET", token=None, data=None):
    headers = {}
    body = None
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if data is not None:
        body = urllib.parse.urlencode(data).encode("utf-8")
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {e.code}: {detail}") from e


def access_token():
    client_id = env("GMAIL_CLIENT_ID")
    client_secret = env("GMAIL_CLIENT_SECRET")
    refresh_token = env("GMAIL_REFRESH_TOKEN")
    if not (client_id and client_secret and refresh_token):
        return None
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token",
    }
    return request_json(TOKEN_URL, method="POST", data=data)["access_token"]


def gmail_query():
    custom = env("GMAIL_NEWS_QUERY")
    if custom:
        return custom
    days = int_env("GMAIL_NEWS_LOOKBACK_DAYS", 14, lo=1, hi=365)
    return f"from:{GOOGLE_ALERTS_FROM} newer_than:{days}d (Ultrahuman OR Wearables)"


def list_messages(token):
    limit = int_env("GMAIL_NEWS_LIMIT", 20, lo=1, hi=100)
    params = urllib.parse.urlencode({"q": gmail_query(), "maxResults": limit})
    data = request_json(f"{GMAIL_API}/messages?{params}", token=token)
    return data.get("messages", [])[:limit]


def get_message(token, msg_id):
    params = urllib.parse.urlencode({"format": "full"})
    return request_json(f"{GMAIL_API}/messages/{msg_id}?{params}", token=token)


def header(msg, name):
    for h in msg.get("payload", {}).get("headers", []):
        if h.get("name", "").lower() == name.lower():
            return h.get("value", "")
    return ""


def b64decode(data):
    if not data:
        return ""
    padding = "=" * (-len(data) % 4)
    return base64.urlsafe_b64decode((data + padding).encode()).decode("utf-8", errors="replace")


def walk_parts(part):
    yield part
    for child in part.get("parts", []) or []:
        yield from walk_parts(child)


def message_text(msg):
    html_parts, text_parts = [], []
    for part in walk_parts(msg.get("payload", {})):
        mime = part.get("mimeType", "")
        data = (part.get("body") or {}).get("data")
        if mime == "text/html":
            html_parts.append(b64decode(data))
        elif mime == "text/plain":
            text_parts.append(b64decode(data))
    raw_html = "\n".join(html_parts)
    raw_text = "\n".join(text_parts)
    return raw_html, raw_text


def clean_text(value):
    value = re.sub(r"<[^>]+>", " ", value or "")
    value = html.unescape(value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def clean_url(url):
    url = html.unescape(url or "")
    parsed = urllib.parse.urlparse(url)
    if "google." in parsed.netloc and parsed.path.startswith("/url"):
        qs = urllib.parse.parse_qs(parsed.query)
        url = (qs.get("q") or qs.get("url") or [url])[0]
    return url.strip()


def domain(url):
    host = urllib.parse.urlparse(url).netloc.lower()
    host = host[4:] if host.startswith("www.") else host
    return host


def message_date(msg):
    raw = header(msg, "Date")
    try:
        return email.utils.parsedate_to_datetime(raw).date().isoformat()
    except Exception:
        ms = int(msg.get("internalDate", "0") or 0)
        if ms:
            return dt.datetime.fromtimestamp(ms / 1000, tz=dt.timezone.utc).date().isoformat()
    return None


def extract_alert_items(msg):
    raw_html, raw_text = message_text(msg)
    date = message_date(msg)
    items = []

    links = re.findall(r'<a\b[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', raw_html, flags=re.I | re.S)
    for href, label_html in links:
        title = clean_text(label_html)
        url = clean_url(href)
        if not title or len(title) < 12:
            continue
        if not url.startswith("http") or "google.com/alerts" in url:
            continue
        hay = f"{title} {raw_text}".lower()
        if "ultrahuman" not in hay and "wearable" not in hay:
            continue
        is_company = "ultrahuman" in hay
        items.append({
            "date": date,
            "type": "Company" if is_company else "Industry",
            "company": "Ultrahuman" if is_company else "",
            "theme": "Wearables",
            "source": domain(url) or "Google Alerts",
            "headline": title,
            "url": url,
            "summary": clean_text(raw_text)[:320],
            "sentiment": "",
            "relevance": "Gmail Alert",
        })

    # Plain-text fallback for alerts whose HTML is sparse.
    if not items and raw_text:
        for url in re.findall(r"https?://\S+", raw_text):
            url = clean_url(url.rstrip(").,]"))
            if "google.com/alerts" in url:
                continue
            text = clean_text(raw_text)
            hay = text.lower()
            if "ultrahuman" not in hay and "wearable" not in hay:
                continue
            items.append({
                "date": date,
                "type": "Company" if "ultrahuman" in hay else "Industry",
                "company": "Ultrahuman" if "ultrahuman" in hay else "",
                "theme": "Wearables",
                "source": domain(url) or "Google Alerts",
                "headline": header(msg, "Subject") or "Google Alert",
                "url": url,
                "summary": text[:320],
                "sentiment": "",
                "relevance": "Gmail Alert",
            })
            break
    return items


def merge_news(data, gmail_items):
    news = data.get("news") or {"items": [], "industry": [], "companies": []}
    existing = news.get("items") or []
    seen = {x.get("url") for x in existing if x.get("url")}
    merged = list(existing)
    for item in gmail_items:
        if item.get("url") in seen:
            continue
        seen.add(item.get("url"))
        merged.append(item)
    merged.sort(key=lambda x: x.get("date") or "", reverse=True)
    news["items"] = merged
    news["industry"] = [x for x in merged if str(x.get("type", "")).lower() == "industry" or not x.get("company")]
    news["companies"] = [x for x in merged if str(x.get("type", "")).lower() == "company" or x.get("company")]
    news["meta"] = {
        "gmailQuery": gmail_query(),
        "gmailItemsAdded": len(merged) - len(existing),
        "gmailItemsSeen": len(gmail_items),
    }
    data["news"] = news


def main():
    out_path = sys.argv[1] if len(sys.argv) > 1 else "data.json"
    data = json.loads(Path(out_path).read_text()) if Path(out_path).exists() else {}
    token = access_token()
    if not token:
        print("! skipped Gmail news: set GMAIL_CLIENT_ID, GMAIL_CLIENT_SECRET and GMAIL_REFRESH_TOKEN")
        return

    items = []
    messages = list_messages(token)
    for msg in messages:
        full = get_message(token, msg["id"])
        items.extend(extract_alert_items(full))

    merge_news(data, items)
    Path(out_path).write_text(json.dumps(data, indent=2, ensure_ascii=False))
    print(f"✓ merged {len(items)} Gmail alert items into news using query: {gmail_query()}")


if __name__ == "__main__":
    main()
