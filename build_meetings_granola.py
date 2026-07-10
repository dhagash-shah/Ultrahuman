#!/usr/bin/env python3
"""
build_meetings_granola.py — pull recent company meeting notes from Granola into
data.json.

Requires GRANOLA_API_KEY. If no key is present, preserves any existing meetings
block and exits without failing the dashboard build.
"""
import datetime
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

DEFAULT_BASE_URL = "https://public-api.granola.ai/v1"
MEETING_TERMS = re.compile(r"\b(meeting|discussion|call|banker|roadmap|aop|board|bod|strategy)\b", re.I)


def api_key():
    return os.environ.get("GRANOLA_API_KEY") or os.environ.get("GRANOLA_TOKEN")


def base_url():
    return (os.environ.get("GRANOLA_API_BASE_URL") or DEFAULT_BASE_URL).rstrip("/")


def int_env(name, default, lo=None, hi=None):
    try:
        value = int(os.environ.get(name, default))
    except ValueError:
        value = default
    if lo is not None:
        value = max(lo, value)
    if hi is not None:
        value = min(hi, value)
    return value


def request_json(path, token, params=None, retries=3):
    url = f"{base_url()}{path}"
    if params:
        clean = {k: v for k, v in params.items() if v not in (None, "")}
        if clean:
            url += "?" + urllib.parse.urlencode(clean)
    req = urllib.request.Request(
        url,
        headers={"Authorization": f"Bearer {token}"},
        method="GET",
    )
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code == 429 and attempt < retries - 1:
                time.sleep(2 ** attempt)
                continue
            body = e.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"Granola API failed HTTP {e.code}: {body}") from e


def created_after_iso():
    raw = os.environ.get("GRANOLA_CREATED_AFTER")
    if raw:
        return raw
    days = int_env("GRANOLA_LOOKBACK_DAYS", 365, lo=1, hi=3650)
    cutoff = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=days)
    return cutoff.replace(microsecond=0).isoformat().replace("+00:00", "Z")


def list_notes(token):
    wanted = int_env("GRANOLA_FETCH_LIMIT", 30, lo=1, hi=120)
    page_size = min(30, wanted)
    params = {
        "created_after": created_after_iso(),
        "folder_id": os.environ.get("GRANOLA_FOLDER_ID"),
        "page_size": page_size,
    }
    notes, cursor = [], None
    while len(notes) < wanted:
        if cursor:
            params["cursor"] = cursor
        data = request_json("/notes", token, params)
        notes.extend(data.get("notes", []))
        if not data.get("hasMore"):
            break
        cursor = data.get("cursor")
        if not cursor:
            break
    return notes[:wanted]


def get_note(token, note_id):
    include = "transcript" if truthy("GRANOLA_INCLUDE_TRANSCRIPT") else None
    return request_json(f"/notes/{note_id}", token, {"include": include})


def truthy(name):
    return str(os.environ.get(name, "")).strip().lower() in {"1", "true", "yes", "y"}


def plain_date(value):
    if not value:
        return None
    return str(value).split("T", 1)[0]


def note_text(note):
    parts = [
        note.get("title"),
        note.get("summary_text"),
        note.get("summary_markdown"),
        ((note.get("calendar_event") or {}).get("event_title")),
    ]
    return "\n".join(str(p) for p in parts if p)


def query_terms():
    raw = os.environ.get("GRANOLA_MEETING_QUERY", "Ultrahuman")
    return [q.strip().lower() for q in raw.split(",") if q.strip()]


def is_wanted(note):
    queries = query_terms()
    text = note_text(note).lower()
    if queries and not any(q in text for q in queries):
        return False
    title = str(note.get("title") or "")
    event_title = str((note.get("calendar_event") or {}).get("event_title") or "")
    return bool(MEETING_TERMS.search(title) or MEETING_TERMS.search(event_title) or queries)


def markdown_bullets(markdown, fallback_text):
    bullets = []
    for line in (markdown or "").splitlines():
        clean = line.strip()
        if re.match(r"^[-*•]\s+", clean):
            clean = re.sub(r"^[-*•]\s+", "", clean).strip()
        elif re.match(r"^\d+[.)]\s+", clean):
            clean = re.sub(r"^\d+[.)]\s+", "", clean).strip()
        else:
            continue
        clean = re.sub(r"[*_`#]+", "", clean).strip()
        if len(clean) > 8:
            bullets.append(clean)
    if not bullets and fallback_text:
        sentences = re.split(r"(?<=[.!?])\s+", fallback_text.strip())
        bullets = [s.strip() for s in sentences if len(s.strip()) > 20]
    return bullets[:5]


def markdown_sections(markdown):
    sections = []
    for line in (markdown or "").splitlines():
        clean = line.strip()
        if clean.startswith("#"):
            clean = clean.lstrip("#").strip()
            if clean:
                sections.append(clean)
    return sections[:4]


def attendees(note):
    people = []
    for person in note.get("attendees") or []:
        name = person.get("name") or person.get("email")
        if name:
            people.append(name)
    return people[:6]


def to_item(note):
    event = note.get("calendar_event") or {}
    markdown = note.get("summary_markdown") or ""
    summary_text = note.get("summary_text") or ""
    return {
        "title": note.get("title") or event.get("event_title") or "Untitled meeting",
        "date": plain_date(event.get("scheduled_start_time") or note.get("created_at")),
        "lastEdited": note.get("updated_at"),
        "url": note.get("web_url"),
        "source": "Granola",
        "sections": markdown_sections(markdown),
        "bullets": markdown_bullets(markdown, summary_text),
        "attendees": attendees(note),
    }


def search_meetings(token):
    wanted = int_env("GRANOLA_MEETING_LIMIT", 5, lo=1, hi=12)
    items = []
    fetched = list_notes(token)
    checked = 0
    fallback = []
    for raw in fetched:
        try:
            note = get_note(token, raw["id"])
        except Exception as e:
            print(f"! skipped Granola note {raw.get('id')}: {e}")
            continue
        checked += 1
        if is_wanted(note):
            items.append(to_item(note))
        elif MEETING_TERMS.search(note_text(note)):
            fallback.append(to_item(note))
        if len(items) >= wanted:
            break
    if not items and truthy("GRANOLA_ALLOW_RECENT_FALLBACK"):
        items = fallback[:wanted]
    items.sort(key=lambda x: x.get("date") or x.get("lastEdited") or "", reverse=True)
    meta = {
        "status": "ok" if items else "empty",
        "query": ", ".join(query_terms()),
        "fetched": len(fetched),
        "checked": checked,
        "matched": len(items),
        "message": None,
    }
    if not items:
        if not fetched:
            meta["message"] = "Granola returned no notes for the configured lookback window."
        else:
            meta["message"] = f"Granola returned {len(fetched)} notes, but none matched: {meta['query'] or 'no query'}."
    return items[:wanted], meta


def main():
    out_path = sys.argv[1] if len(sys.argv) > 1 else "data.json"
    data = json.loads(Path(out_path).read_text()) if Path(out_path).exists() else {}
    token = api_key()
    if not token:
        data["meetings"] = {
            "source": "Granola",
            "items": [],
            "meta": {
                "status": "missing_api_key",
                "message": "Set GRANOLA_API_KEY in this GitHub repository.",
            },
        }
        Path(out_path).write_text(json.dumps(data, indent=2, ensure_ascii=False))
        print("! skipped Granola meetings: set GRANOLA_API_KEY")
        return
    try:
        items, meta = search_meetings(token)
    except Exception as e:
        sys.exit(f"ERROR: {e}")
    data["meetings"] = {"source": "Granola", "items": items, "meta": meta}
    Path(out_path).write_text(json.dumps(data, indent=2, ensure_ascii=False))
    print(f"✓ wrote {len(items)} recent Granola meetings to {out_path}")
    if meta.get("message"):
        print(f"  {meta['message']}")


if __name__ == "__main__":
    main()
