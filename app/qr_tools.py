from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Any

import qrcode


@dataclass
class Attendee:
    id: str
    name: str
    extra: Dict[str, Any]


def load_attendees_csv(csv_path: str | Path) -> List[Attendee]:
    path = Path(csv_path)
    rows: List[Attendee] = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        headers = [h.strip() for h in (reader.fieldnames or [])]
        if "id" not in headers or "name" not in headers:
            raise ValueError("CSV 必須包含欄位 id, name")
        for r in reader:
            rid = str(r.get("id", "")).strip()
            nm = str(r.get("name", "")).strip()
            extra = {k: v for k, v in r.items() if k not in ("id", "name") and v is not None and str(v).strip() != ""}
            if rid or nm:
                rows.append(Attendee(id=rid, name=nm, extra=extra))
    return rows


def export_template_csv(path: str | Path, extra_fields: list[str] | None = None) -> None:
    path = Path(path)
    extras = [e for e in (extra_fields or ["email", "company"]) if e not in ("id", "name")]
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        # 只輸出標題列；不包含示例資料列
        writer.writerow(["id", "name", *extras])


def generate_qr_images(attendees: List[Attendee], event_name: str, out_dir: str | Path) -> int:
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    count = 0
    for a in attendees:
        payload = {
            "id": a.id,
            "name": a.name,
            "event": event_name,
            "extra": a.extra,
        }
        data = json.dumps(payload, ensure_ascii=False)
        img = qrcode.make(data)
        fn = f"{a.id or a.name}.png".replace("/", "_").replace("\\", "_")
        img.save(out / fn)
        count += 1
    return count


def parse_qr_payload(data: str, default_event: str) -> Dict[str, Any]:
    try:
        obj = json.loads(data)
        if isinstance(obj, dict):
            obj.setdefault("event", default_event)
            return obj
    except Exception:
        pass
    # Fallback: treat as raw
    return {"raw": data, "event": default_event}
