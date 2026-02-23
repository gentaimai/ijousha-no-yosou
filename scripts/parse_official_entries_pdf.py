#!/usr/bin/env python3
"""Parse the 2025 Japan Championships entry PDF into app master-data JSON.

Usage:
  python scripts/parse_official_entries_pdf.py /path/to/entries.pdf

Requires:
  pypdf
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

from pypdf import PdfReader


HEADER_SKIP = {
    "氏名: カナ: 所属: 学校:人数:",
    "第101回日本選手権水泳競技大会",
}
FOOTER_RE = re.compile(r"^\d+/\d+\s+ページ")
ROW_RE = re.compile(r"^(?P<body>.+?)(?P<seed>\d+)\s+(?P<code>[0-9A-Z]{5})$")
EVENT_RE = re.compile(r"^(男子|女子)\s*(.+?)\s*$")
NAME_RE = re.compile(r"^(?P<name>.+?)\s+[ｦ-ﾟ]")


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("\u3000", " ")).strip()


def parse_pdf(pdf_path: Path) -> dict:
    reader = PdfReader(str(pdf_path))
    events = []
    entries = []
    event_map = {}

    for page in reader.pages:
        text = page.extract_text() or ""
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if not lines:
            continue

        m_event = EVENT_RE.match(lines[0])
        if not m_event:
            continue

        gender = m_event.group(1)
        event_name = normalize_space(m_event.group(2))
        event_key = f"{gender}:{event_name}"
        if event_key not in event_map:
            event_id = f"{'M' if gender == '男子' else 'W'}-{len(events) + 1:02d}"
            event = {
                "eventId": event_id,
                "sortOrder": len(events) + 1,
                "gender": gender,
                "eventName": event_name,
            }
            event_map[event_key] = event
            events.append(event)

        event_id = event_map[event_key]["eventId"]

        for raw in lines[1:]:
            if raw in HEADER_SKIP or FOOTER_RE.match(raw):
                continue
            m_row = ROW_RE.match(raw)
            if not m_row:
                continue

            body = normalize_space(m_row.group("body"))
            m_name = NAME_RE.match(body)
            if not m_name:
                continue

            athlete_name = normalize_space(m_name.group("name"))
            seed = int(m_row.group("seed"))
            org_code = m_row.group("code")

            entries.append(
                {
                    "entryId": f"{event_id}:{seed:03d}:{org_code}",
                    "eventId": event_id,
                    "seedOrder": seed,
                    "athleteName": athlete_name,
                    "team": "",
                    "entryTime": "",
                }
            )

    entries.sort(key=lambda r: (r["eventId"], r["seedOrder"], r["entryId"]))
    return {"events": events, "entries": entries}


def main() -> int:
    if len(sys.argv) != 2:
        print("usage: python scripts/parse_official_entries_pdf.py /path/to/entries.pdf", file=sys.stderr)
        return 2

    payload = parse_pdf(Path(sys.argv[1]))
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

