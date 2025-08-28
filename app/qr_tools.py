from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Any, Tuple, Optional

import qrcode
from PIL import Image, ImageDraw, ImageFont


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
    extras = [e for e in (extra_fields or ["salon", "seller"]) if e not in ("id", "name")]
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


# ===== Designed Poster Rendering =====
@dataclass
class DesignOptions:
    width: int = 1080
    height: int = 1350
    qr_ratio: float = 0.7  # QR width relative to canvas width (0.0~1.0)
    bg_color: str = "#FFFFFF"
    qr_color: str = "#000000"
    text_color: str = "#000000"
    font_family: Optional[str] = None
    font_path: Optional[str] = None  # optional direct path override
    font_size: int = 48
    margin: int = 40
    bg_image_path: Optional[str] = None
    text_anchor: str = "top"  # default top within region
    text_align: str = "center"   # one of: left, center, right
    text_margin: int = 40         # inner margin for text block
    line_spacing_scale: float = 0.4
    auto_fit_text: bool = True    # shrink font to fit region if needed
    text_region: Optional[Tuple[float, float, float, float]] = None  # (x0,y0,x1,y1) normalized


def _hex_to_rgb(hex_str: str) -> Tuple[int, int, int]:
    s = hex_str.strip()
    if s.startswith('#'):
        s = s[1:]
    if len(s) == 3:
        s = ''.join(ch * 2 for ch in s)
    try:
        r = int(s[0:2], 16)
        g = int(s[2:4], 16)
        b = int(s[4:6], 16)
        return (r, g, b)
    except Exception:
        return (0, 0, 0)


def _find_font_file(family: Optional[str]) -> Optional[str]:
    if not family:
        return None
    try:
        import os
        from pathlib import Path
        candidates: List[Path] = []
        if os.name == 'nt':
            candidates.append(Path(os.environ.get('WINDIR', 'C:/Windows')) / 'Fonts')
        else:
            candidates += [
                Path('/usr/share/fonts'),
                Path('/usr/local/share/fonts'),
                Path.home() / '.fonts',
                Path('/Library/Fonts'),
                Path('/System/Library/Fonts'),
                Path.home() / 'Library/Fonts',
            ]
        # Direct known-mapping for common CJK on Windows
        known_map = {
            'microsoft jhenghei': 'msjh.ttc',
            '微軟正黑體': 'msjh.ttc',
            'mingliu': 'mingliu.ttc',
            'pmingliu': 'pmingliu.ttc',
            '新細明體': 'mingliu.ttc',
            '細明體': 'mingliu.ttc',
            '標楷體': 'kaiu.ttf',
            'kaiu': 'kaiu.ttf',
            'noto sans cjk': 'NotoSansCJK-Regular.ttc',
        }
        family_norm = family.lower().strip()
        if os.name == 'nt' and family_norm in known_map:
            for root in candidates:
                p = root / known_map[family_norm]
                if p.exists():
                    return str(p)
        # Fuzzy match by filename stem
        family_norm2 = family_norm.replace(' ', '').replace('-', '').replace('_', '')
        for root in candidates:
            if not root.exists():
                continue
            for p in root.rglob('*'):
                if p.suffix.lower() not in ('.ttf', '.otf', '.ttc'):
                    continue
                name = p.stem.lower().replace(' ', '').replace('-', '').replace('_', '')
                if family_norm2 in name:
                    return str(p)
    except Exception:
        return None
    return None


def get_font_with_meta(size: int, family: Optional[str] = None, path: Optional[str] = None):
    """Return (font, meta) where meta includes resolved path, index, and family name.
    Tries hard to map a family to an actual font file and TTC face index.
    """
    try:
        fpath = path or _find_font_file(family)
        # last-resort common CJK candidates
        if not fpath:
            for c in [
                'C:/Windows/Fonts/msjh.ttc',  # 微軟正黑體
                'C:/Windows/Fonts/msyh.ttc',  # 微軟雅黑
                'C:/Windows/Fonts/mingliu.ttc',
                'C:/Windows/Fonts/simhei.ttf',
                'C:/Windows/Fonts/simsun.ttc',
                '/System/Library/Fonts/PingFang.ttc',
                '/Library/Fonts/Arial Unicode.ttf',
                '/Library/Fonts/Arial Unicode MS.ttf',
                '/System/Library/Fonts/STHeiti Light.ttc',
                '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
                '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
            ]:
                if Path(c).exists():
                    fpath = c
                    break
        if fpath and Path(fpath).exists():
            fpl = str(fpath).lower()
            if fpl.endswith('.ttc') and family:
                # Search faces to match requested family
                family_norm = family.lower().replace(' ', '')
                first: Optional[tuple] = None
                for idx in range(0, 8):
                    try:
                        ft = ImageFont.truetype(fpath, size, index=idx)
                        fam = ''
                        try:
                            fam = ft.getname()[0]
                        except Exception:
                            fam = ''
                        if family_norm in fam.lower().replace(' ', ''):
                            return ft, {"path": str(fpath), "index": idx, "name": fam}
                        if first is None:
                            first = (ft, {"path": str(fpath), "index": idx, "name": fam})
                    except Exception:
                        continue
                if first is not None:
                    return first
            # Non-TTC or no family match
            ft = ImageFont.truetype(fpath, size)
            fam = ''
            try:
                fam = ft.getname()[0]
            except Exception:
                fam = ''
            return ft, {"path": str(fpath), "index": 0, "name": fam}
    except Exception:
        return ImageFont.load_default(), {"path": "default", "index": 0, "name": "default"}


def _get_font(size: int, family: Optional[str] = None, path: Optional[str] = None) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    try:
        f, _meta = get_font_with_meta(size, family, path)
        return f
    except Exception:
        try:
            return ImageFont.load_default()
        except Exception:
            return ImageFont.load_default()


def generate_qr_posters(attendees: List[Attendee], event_name: str, out_dir: str | Path, options: DesignOptions) -> int:
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    count = 0
    for a in attendees:
        # Build QR payload JSON
        payload = {
            "id": a.id,
            "name": a.name,
            "event": event_name,
            "extra": a.extra,
        }
        data = json.dumps(payload, ensure_ascii=False)

        # Create QR image
        qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M)
        qr.add_data(data)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color=options.qr_color, back_color="#FFFFFF").convert("RGB")

        # Create canvas (image or solid color)
        if options.bg_image_path and Path(options.bg_image_path).exists():
            bg = Image.open(options.bg_image_path).convert("RGB")
            # cover resize
            sw, sh = bg.size
            tw, th = options.width, options.height
            scale = max(tw / sw, th / sh)
            new_size = (int(sw * scale), int(sh * scale))
            bg = bg.resize(new_size, Image.Resampling.LANCZOS)
            # center crop
            left = (bg.width - tw) // 2
            top = (bg.height - th) // 2
            canvas = bg.crop((left, top, left + tw, top + th))
        else:
            canvas = Image.new("RGB", (options.width, options.height), _hex_to_rgb(options.bg_color))
        draw = ImageDraw.Draw(canvas)

        # Compute QR size and position (centered horizontally, placed upper area)
        qr_target_w = int(options.width * max(0.1, min(1.0, options.qr_ratio)))
        qr_target_w = max(50, min(qr_target_w, options.width - options.margin * 2))
        # Resize QR to square target
        qr_resized = qr_img.resize((qr_target_w, qr_target_w), Image.Resampling.LANCZOS)
        qr_x = (options.width - qr_target_w) // 2
        qr_y = options.margin
        canvas.paste(qr_resized, (qr_x, qr_y))

        # Prepare font
        font = _get_font(options.font_size, options.font_family, options.font_path)
        text_color = _hex_to_rgb(options.text_color)

        # Compose lines: id, name, salon, seller (from extra)
        salon = str(a.extra.get("salon", ""))
        seller = str(a.extra.get("seller", ""))
        lines = [
            f"ID: {a.id}",
            f"姓名: {a.name}",
            f"Salon: {salon}",
            f"Seller: {seller}",
        ]

        # Draw text block in area below the QR (qr_bottom .. canvas_bottom)
        line_metrics = []
        total_h = 0
        for line in lines:
            if not line.strip():
                line_metrics.append((line, 0, 0))
                continue
            bbox = draw.textbbox((0, 0), line, font=font)
            w = bbox[2] - bbox[0]
            h = bbox[3] - bbox[1]
            line_metrics.append((line, w, h))
            total_h += h
        non_empty = [m for m in line_metrics if m[2] > 0]
        max_h = max((m[2] for m in non_empty), default=options.font_size)
        spacing = int(max(0.0, options.line_spacing_scale) * max_h)
        if non_empty:
            total_h += spacing * (len(non_empty) - 1)
        # anchor within region: explicit region if provided, else below-QR block
        margin = max(0, int(options.text_margin))
        qr_bottom = qr_y + qr_target_w
        if options.text_region:
            x0n, y0n, x1n, y1n = options.text_region
            region_left = int(max(0.0, min(1.0, x0n)) * options.width)
            region_right = int(max(0.0, min(1.0, x1n)) * options.width)
            region_top = int(max(0.0, min(1.0, y0n)) * options.height)
            region_bottom = int(max(0.0, min(1.0, y1n)) * options.height)
            if region_right < region_left:
                region_left, region_right = region_right, region_left
            if region_bottom < region_top:
                region_top, region_bottom = region_bottom, region_top
        else:
            region_left = margin
            region_right = options.width - margin
            region_top = qr_bottom + margin
            region_bottom = options.height - margin
        region_height = max(0, region_bottom - region_top)
        # optional auto-fit text: shrink font to fit region height
        if options.auto_fit_text and total_h > region_height and region_height > 0:
            scale = max(0.1, region_height / total_h)
            new_size = max(10, int(options.font_size * scale))
            if new_size != options.font_size:
                font = _get_font(new_size, options.font_family, options.font_path)
                # recompute metrics with new font
                line_metrics = []
                total_h = 0
                for line in lines:
                    if not line.strip():
                        line_metrics.append((line, 0, 0))
                        continue
                    bbox = draw.textbbox((0, 0), line, font=font)
                    w = bbox[2] - bbox[0]
                    h = bbox[3] - bbox[1]
                    line_metrics.append((line, w, h))
                    total_h += h
                non_empty = [m for m in line_metrics if m[2] > 0]
                if non_empty:
                    max_h = max((m[2] for m in non_empty), default=new_size)
                    spacing = int(max(0.0, options.line_spacing_scale) * max_h)
                    total_h += spacing * (len(non_empty) - 1)
        # default fallback if region is too small: start at region_top
        if options.text_anchor == "top":
            y = region_top
        elif options.text_anchor == "middle":
            y = region_top + max(0, (region_height - total_h) // 2)
        else:  # bottom
            y = max(region_top, region_bottom - total_h)

        for line, w, h in line_metrics:
            if h == 0:
                continue
            # alignment
            if options.text_align == "left":
                x = region_left
            elif options.text_align == "right":
                x = max(region_left, region_right - w)
            else:
                x = (region_left + region_right - w) // 2

            if y + h > region_bottom:
                break
            draw.text((x, y), line, font=font, fill=text_color)
            y += h + spacing

        fn = f"{a.id or a.name}.png".replace("/", "_").replace("\\", "_")
        canvas.save(out / fn, format="PNG")
        count += 1
    return count
