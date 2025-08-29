from __future__ import annotations

import csv
import json
from typing import List, Dict
from pathlib import Path

from .paths import offline_queue_file, ensure_dirs


FIELD = "payload_json"


def append_payload(payload: Dict) -> None:
    ensure_dirs()
    path = offline_queue_file()
    is_new = not path.exists()
    with path.open("a", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[FIELD])
        if is_new:
            writer.writeheader()
        writer.writerow({FIELD: json.dumps(payload, ensure_ascii=False)})


def read_payloads() -> List[Dict]:
    path = offline_queue_file()
    items: List[Dict] = []
    if not path.exists():
        return items
    with path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            try:
                raw = row.get(FIELD, "")
                if raw:
                    items.append(json.loads(raw))
            except Exception:
                # skip broken rows
                continue
    return items


def write_payloads(payloads: List[Dict]) -> None:
    ensure_dirs()
    path = offline_queue_file()
    if not payloads:
        try:
            if path.exists():
                path.unlink()
        except Exception:
            pass
        return
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[FIELD])
        writer.writeheader()
        for p in payloads:
            writer.writerow({FIELD: json.dumps(p, ensure_ascii=False)})

