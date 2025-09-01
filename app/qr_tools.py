from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Any, Tuple, Optional

import qrcode
from PIL import Image, ImageDraw, ImageFont
import os
if os.name == 'nt':
    try:
        import winreg  # type: ignore
    except Exception:
        winreg = None  # type: ignore


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
        raw_headers = [h for h in (reader.fieldnames or [])]
        headers = [str(h).strip() for h in raw_headers]
        # case-insensitive mapping for required columns
        low_map = {str(h).strip().lower(): str(h).strip() for h in raw_headers if h is not None}
        id_key = low_map.get("id")
        name_key = low_map.get("name")
        if not id_key or not name_key:
            raise ValueError("CSV 必須包含欄位 id, name（不分大小寫）")
        for r in reader:
            rid = str(r.get(id_key, "")).strip()
            nm = str(r.get(name_key, "")).strip()
            extra = {k: v for k, v in r.items() if k not in (id_key, name_key) and v is not None and str(v).strip() != ""}
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
    # font styling
    font_weight: Optional[str] = None  # e.g. 'regular','medium','semibold','bold','extrabold','black'
    font_bold: bool = False
    font_italic: bool = False
    font_underline: bool = False
    margin: int = 40
    bg_image_path: Optional[str] = None
    text_anchor: str = "top"  # default top within region
    text_align: str = "center"   # one of: left, center, right
    text_margin: int = 40         # inner margin for text block
    text_top_gap: int = 40        # vertical gap between QR bottom and first text line
    text_bottom_margin: int = 40  # bottom padding to avoid clipping
    line_spacing_scale: float = 0.4
    auto_fit_text: bool = True    # shrink font to fit region if needed
    text_region: Optional[Tuple[float, float, float, float]] = None  # (x0,y0,x1,y1) normalized
    text_point: Optional[Tuple[float, float]] = None  # normalized (x,y) text position; overrides text_region if set


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


def get_font_with_meta(size: int,
                       family: Optional[str] = None,
                       path: Optional[str] = None,
                       weight: Optional[str] = None,
                       italic: bool = False,
                       bold: bool = False):
    """Return (font, meta) where meta includes resolved path, index, and family name.
    Tries hard to map a family to an actual font file and TTC face index.
    """
    try:
        # On Windows, consult font registry to resolve a more precise variant
        fpath = None
        if path:
            fpath = path
        elif os.name == 'nt' and family and winreg is not None:
            try:
                def _win_lookup(fam: str) -> Optional[str]:
                    base_dir = Path(os.environ.get('WINDIR', 'C:/Windows')) / 'Fonts'
                    best_path: Optional[Path] = None
                    best_score = -1
                    fam_l = fam.lower().strip()
                    want_bold = bold or (isinstance(weight, str) and 'bold' in weight.lower())
                    weight_kw = {
                        'regular': ['regular', 'book', 'roman', 'normal'],
                        'medium': ['medium'],
                        'semibold': ['semibold', 'demibold', 'semi bold', 'demi bold'],
                        'bold': ['bold'],
                        'extrabold': ['extrabold', 'extra bold', 'heavy', 'black'],
                        'black': ['black', 'heavy', 'ultra'],
                    }
                    want_set = set()
                    if isinstance(weight, str):
                        want_set |= set(weight_kw.get(weight.lower(), []))
                    if want_bold:
                        want_set |= set(weight_kw.get('bold', []))
                    if italic:
                        want_set |= {'italic', 'oblique'}
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") as k:
                            i = 0
                            while True:
                                try:
                                    name, value, _ = winreg.EnumValue(k, i)
                                except OSError:
                                    break
                                i += 1
                                name_l = str(name).lower()
                                # must roughly match family
                                fam_hit = fam_l in name_l
                                if not fam_hit:
                                    continue
                                # score by style keywords
                                sc = 0
                                for w in want_set:
                                    if w in name_l:
                                        sc += 2
                                if fam_hit:
                                    sc += 1
                                # prefer italic exact if requested
                                if italic and ('italic' in name_l or 'oblique' in name_l):
                                    sc += 1
                                # compute candidate path
                                p = Path(value)
                                if not p.is_absolute():
                                    p = base_dir / p
                                if p.exists() and sc > best_score:
                                    best_score = sc
                                    best_path = p
                    except Exception:
                        pass
                    return str(best_path) if best_path is not None else None
                fpath = _win_lookup(family)
            except Exception:
                fpath = None
        if not fpath:
            fpath = _find_font_file(family)
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
                best: Optional[tuple] = None
                want_bold = bold or (isinstance(weight, str) and 'bold' in weight.lower())
                # allow mapping of common weight names
                def _style_score(style_name: str) -> int:
                    s = style_name.lower()
                    score = 0
                    if want_bold and ('bold' in s or 'black' in s or 'heavy' in s or 'semi' in s):
                        score += 2
                    if italic and ('italic' in s or 'oblique' in s):
                        score += 2
                    # extra hints for weight text
                    if isinstance(weight, str):
                        w = weight.lower()
                        if w in s:
                            score += 1
                    return score
                for idx in range(0, 32):
                    try:
                        ft = ImageFont.truetype(fpath, size, index=idx)
                        fam = ''
                        try:
                            fam, subfam = ft.getname()
                        except Exception:
                            fam, subfam = '', ''
                        if family_norm in fam.lower().replace(' ', ''):
                            # exact family; record first
                            if first is None:
                                first = (ft, {"path": str(fpath), "index": idx, "name": fam, "style": subfam})
                            # prefer matching style
                            sc = _style_score(subfam)
                            if sc > 0:
                                best = (ft, {"path": str(fpath), "index": idx, "name": fam, "style": subfam})
                                if sc >= 3:
                                    break
                        if first is None:
                            first = (ft, {"path": str(fpath), "index": idx, "name": fam, "style": subfam})
                    except Exception:
                        continue
                if best is not None:
                    return best
                if first is not None:
                    return first
            # Non-TTC or no family match
            p0 = Path(fpath)
            want_bold = bold or (isinstance(weight, str) and 'bold' in weight.lower())
            def _pick_variant(base: Path) -> Path:
                try:
                    base_key = (family or base.stem).lower().replace(' ', '').replace('-', '').replace('_', '')
                    candidates = [q for q in base.parent.glob('*') if q.suffix.lower() in ('.ttf','.otf')]
                    def score(q: Path) -> int:
                        s = q.stem.lower()
                        sc = 0
                        if base_key in s:
                            sc += 1
                        if want_bold and ('bold' in s or 'black' in s or 'heavy' in s or 'semi' in s):
                            sc += 2
                        if italic and ('italic' in s or 'oblique' in s):
                            sc += 2
                        if isinstance(weight, str) and weight.lower() in s:
                            sc += 1
                        return sc
                    best = None
                    best_sc = -1
                    for q in candidates:
                        sc = score(q)
                        if sc > best_sc:
                            best_sc = sc
                            best = q
                    return best or base
                except Exception:
                    return base
            if (want_bold or italic or isinstance(weight, str)):
                p1 = _pick_variant(p0)
            else:
                p1 = p0
            ft = ImageFont.truetype(str(p1), size)
            fam = ''
            try:
                fam, subfam = ft.getname()
            except Exception:
                fam, subfam = '', ''
            return ft, {"path": str(p1), "index": 0, "name": fam, "style": subfam}
    except Exception:
        return ImageFont.load_default(), {"path": "default", "index": 0, "name": "default"}


def _get_font(size: int,
              family: Optional[str] = None,
              path: Optional[str] = None,
              weight: Optional[str] = None,
              italic: bool = False,
              bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    try:
        f, _meta = get_font_with_meta(size, family, path, weight=weight, italic=italic, bold=bold)
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

        # Compute QR block: above text anchor, minus top and side margins
        lr_margin = max(0, int(options.text_margin))
        top_bound = max(0, int(getattr(options, 'text_top_gap', 40)))
        if options.text_point is not None:
            _, yn = options.text_point
            anchor_y = int(max(0.0, min(1.0, float(yn))) * options.height)
        else:
            # default anchor Y if not provided
            anchor_y = int(0.8 * options.height)
        region_left = lr_margin
        region_right = max(region_left + 1, options.width - lr_margin)
        region_top = top_bound
        # 底界也套用上邊界：以錨點往上再留出上邊界
        region_bottom = anchor_y - top_bound
        if region_bottom <= region_top:
            region_bottom = region_top + 1
        region_w = max(1, region_right - region_left)
        region_h = max(1, region_bottom - region_top)
        # QR target size by ratio, clamped to block
        qr_target_w = int(options.width * max(0.05, min(1.0, options.qr_ratio)))
        qr_target_w = max(50, min(qr_target_w, region_w, region_h))
        # Resize QR to square target
        qr_resized = qr_img.resize((qr_target_w, qr_target_w), Image.Resampling.LANCZOS)
        # Center QR within the region
        cx = region_left + region_w / 2.0
        cy = region_top + region_h / 2.0
        qr_x = int(round(cx - qr_target_w / 2.0))
        qr_y = int(round(cy - qr_target_w / 2.0))
        canvas.paste(qr_resized, (qr_x, qr_y))

        # Prepare font
        font = _get_font(options.font_size,
                         options.font_family,
                         options.font_path,
                         weight=options.font_weight)
        text_color = _hex_to_rgb(options.text_color)

        # Compose lines: id, name, salon, seller (from extra)
        # Be tolerant to CSV header casing (e.g., "Salon" vs "salon")
        def _extra_ci(extra: Dict[str, Any], key: str) -> str:
            try:
                k_l = key.lower()
                for k, v in extra.items():
                    if str(k).lower() == k_l:
                        return str(v)
            except Exception:
                pass
            return ""
        salon = _extra_ci(a.extra, "salon")
        seller = _extra_ci(a.extra, "seller")
        # keep label casing based on CSV headers for extras when possible
        def _label_from_extra(extra: Dict[str, Any], key: str, default: str) -> str:
            try:
                k_l = key.lower()
                for k in extra.keys():
                    if str(k).lower() == k_l:
                        return str(k)
            except Exception:
                pass
            return default
        salon_label = _label_from_extra(a.extra, "salon", "salon")
        seller_label = _label_from_extra(a.extra, "seller", "seller")
        lines = [
            f"ID: {a.id}",
            f"name: {a.name}",
            f"{salon_label}: {salon}",
            f"{seller_label}: {seller}",
        ]

        # Try Qt text rendering for better font support; fallback to PIL if Qt not available
        try:
            from PyQt6 import QtGui, QtCore  # type: ignore
            from PIL.ImageQt import ImageQt  # type: ignore
            _qt_ok = True
        except Exception:
            _qt_ok = False
        if _qt_ok:
            qimg = ImageQt(canvas).copy()
            painter = QtGui.QPainter(qimg)
            painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)
            # Prepare QFont from options
            qf = QtGui.QFont()
            if options.font_family:
                qf.setFamily(options.font_family)
            qf.setPixelSize(int(options.font_size))
            wmap = {
                'thin': QtGui.QFont.Weight.Thin,
                'extralight': QtGui.QFont.Weight.ExtraLight,
                'light': QtGui.QFont.Weight.Light,
                'regular': QtGui.QFont.Weight.Normal,
                'medium': QtGui.QFont.Weight.Medium,
                'semibold': QtGui.QFont.Weight.DemiBold,
                'bold': QtGui.QFont.Weight.Bold,
                'extrabold': QtGui.QFont.Weight.ExtraBold,
                'black': QtGui.QFont.Weight.Black,
            }
            if getattr(options, 'font_weight', None):
                qf.setWeight(wmap.get(str(options.font_weight).lower(), QtGui.QFont.Weight.Normal))
            # BIU removed; only weight is applied
            painter.setFont(qf)
            # Color
            qc = QtGui.QColor(options.text_color)
            if not qc.isValid():
                qc = QtGui.QColor('#000000')
            painter.setPen(qc)
            fm = QtGui.QFontMetricsF(qf)
            rows = []  # (lab, val, wL, wR, h)
            total_h = 0
            max_h = 0
            for line in lines:
                parts = line.split(":", 1)
                lab = parts[0].strip()
                val = parts[1].strip() if len(parts) > 1 else ""
                brL = fm.boundingRect(lab)
                brV = fm.boundingRect(val)
                wL = int(brL.width()); hL = int(fm.height())
                wR = int(brV.width()); hR = int(fm.height())
                h = max(hL, hR)
                rows.append((lab, val, wL, wR, h))
                total_h += h
                max_h = max(max_h, h)
            spacing = int(max(0.0, options.line_spacing_scale) * (max_h or options.font_size))
            if rows:
                total_h += spacing * (len(rows) - 1)
            # text_margin: only controls left/right padding (not vertical)
            lr_margin = max(0, int(options.text_margin))
            top_gap = max(0, int(getattr(options, 'text_top_gap', 40)))
            bottom_margin = max(0, int(getattr(options, 'text_bottom_margin', 40)))
            region_bottom = options.height - bottom_margin
            # Origin from point (Y) or below-QR region
            if options.text_point is not None:
                _, yn = options.text_point
                y = int(max(0.0, min(1.0, yn)) * options.height)
                # auto-fit in point-anchored mode as well (match preview)
                avail_h = max(0, region_bottom - y)
                if options.auto_fit_text and total_h > avail_h and avail_h > 0:
                    scale = max(0.1, avail_h / max(1, total_h))
                    new_px = max(10, int(options.font_size * scale))
                    if new_px != options.font_size:
                        qf.setPixelSize(new_px)
                        painter.setFont(qf)
                        fm = QtGui.QFontMetricsF(qf)
                        rows = []
                        total_h = 0
                        max_h = 0
                        for line in lines:
                            parts = line.split(":", 1)
                            lab = parts[0].strip()
                            val = parts[1].strip() if len(parts) > 1 else ""
                            brL = fm.boundingRect(lab)
                            brV = fm.boundingRect(val)
                            wL = int(brL.width()); hL = int(fm.height())
                            wR = int(brV.width()); hR = int(fm.height())
                            h = max(hL, hR)
                            rows.append((lab, val, wL, wR, h))
                            total_h += h
                            max_h = max(max_h, h)
                        spacing = int(max(0.0, options.line_spacing_scale) * (max_h or new_px))
                        if rows:
                            total_h += spacing * (len(rows) - 1)
            else:
                region_top = (qr_y + qr_target_w) + top_gap
                region_height = max(0, region_bottom - region_top)
                if options.auto_fit_text and total_h > region_height and region_height > 0:
                    scale = max(0.1, region_height / total_h)
                    new_px = max(10, int(options.font_size * scale))
                    if new_px != options.font_size:
                        qf.setPixelSize(new_px)
                        painter.setFont(qf)
                        fm = QtGui.QFontMetricsF(qf)
                        rows = []
                        total_h = 0
                        max_h = 0
                        for line in lines:
                            parts = line.split(":", 1)
                            lab = parts[0].strip()
                            val = parts[1].strip() if len(parts) > 1 else ""
                            brL = fm.boundingRect(lab)
                            brV = fm.boundingRect(val)
                            wL = int(brL.width()); hL = int(fm.height())
                            wR = int(brV.width()); hR = int(fm.height())
                            h = max(hL, hR)
                            rows.append((lab, val, wL, wR, h))
                            total_h += h
                            max_h = max(max_h, h)
                        spacing = int(max(0.0, options.line_spacing_scale) * (max_h or new_px))
                        if rows:
                            total_h += spacing * (len(rows) - 1)
                if options.text_anchor == 'middle':
                    y = region_top + max(0, (region_bottom - region_top - total_h) // 2)
                elif options.text_anchor == 'bottom':
                    y = max(region_top, region_bottom - total_h)
                else:
                    y = region_top
            # Draw two columns (Qt baseline): header left, content right, with eliding
            cur_y = y + int(fm.ascent())
            region_left = lr_margin
            region_right = options.width - lr_margin
            total_w = max(0, region_right - region_left)
            min_gap = int(max(8, options.font_size * 0.40))
            for lab, val, wL, wR, h in rows:
                if h == 0:
                    continue
                if cur_y + (h - int(fm.ascent())) > region_bottom:
                    break
                max_val_w = max(0, total_w - min_gap)
                val_draw = fm.elidedText(val, QtCore.Qt.TextElideMode.ElideRight, max_val_w)
                wR2 = int(fm.boundingRect(val_draw).width())
                avail_left = max(0, total_w - min_gap - wR2)
                lab_draw = fm.elidedText(lab, QtCore.Qt.TextElideMode.ElideRight, avail_left)
                wL2 = int(fm.boundingRect(lab_draw).width())
                xL = region_left
                xR = region_right - wR2
                painter.drawText(QtCore.QPointF(float(xL), float(cur_y)), lab_draw)
                painter.drawText(QtCore.QPointF(float(xR), float(cur_y)), val_draw)
                cur_y += h + spacing
            painter.end()
            fn = f"{a.id or a.name}.png".replace("/", "_").replace("\\", "_")
            qimg.save(str(out / fn), "PNG")
            count += 1
            continue

        # Draw text block (fallback PIL): two-column layout
        line_metrics = []  # (lab, val, wL, wR, h)
        total_h = 0
        for line in lines:
            if not line.strip():
                line_metrics.append(("", "", 0, 0, 0))
                continue
            parts = line.split(":", 1)
            lab = parts[0].strip()
            val = parts[1].strip() if len(parts) > 1 else ""
            bboxL = draw.textbbox((0, 0), lab, font=font)
            bboxR = draw.textbbox((0, 0), val, font=font)
            wL = bboxL[2] - bboxL[0]
            hL = bboxL[3] - bboxL[1]
            wR = bboxR[2] - bboxR[0]
            hR = bboxR[3] - bboxR[1]
            h = max(hL, hR)
            line_metrics.append((lab, val, wL, wR, h))
            total_h += h
        non_empty = [m for m in line_metrics if m[4] > 0]
        max_h = max((m[4] for m in non_empty), default=options.font_size)
        spacing = int(max(0.0, options.line_spacing_scale) * max_h)
        if non_empty:
            total_h += spacing * (len(non_empty) - 1)
        # text_margin: only controls left/right padding (not vertical)
        lr_margin = max(0, int(options.text_margin))
        top_gap = max(0, int(getattr(options, 'text_top_gap', 40)))
        bottom_margin = max(0, int(getattr(options, 'text_bottom_margin', 40)))
        # If a direct text point is provided, use it; else fall back to region logic
        if options.text_point is not None:
            xn, yn = options.text_point
            x_base = int(max(0.0, min(1.0, xn)) * options.width)
            y_top = int(max(0.0, min(1.0, yn)) * options.height)
            # auto-fit against space to the bottom of canvas (respect bottom margin)
            region_bottom = options.height - bottom_margin
            avail_h = max(0, region_bottom - y_top)
            if options.auto_fit_text and total_h > avail_h and avail_h > 0:
                scale = max(0.1, avail_h / max(1, total_h))
                new_size = max(10, int(options.font_size * scale))
                if new_size != options.font_size:
                    font = _get_font(new_size, options.font_family, options.font_path)
                    # recompute metrics
                    line_metrics = []
                    total_h = 0
            for line in lines:
                if not line.strip():
                    line_metrics.append(("", "", 0, 0, 0))
                    continue
                parts = line.split(":", 1)
                lab = parts[0].strip()
                val = parts[1].strip() if len(parts) > 1 else ""
                bboxL = draw.textbbox((0, 0), lab, font=font)
                bboxR = draw.textbbox((0, 0), val, font=font)
                wL = bboxL[2] - bboxL[0]
                hL = bboxL[3] - bboxL[1]
                wR = bboxR[2] - bboxR[0]
                hR = bboxR[3] - bboxR[1]
                h = max(hL, hR)
                line_metrics.append((lab, val, wL, wR, h))
                total_h += h
            non_empty = [m for m in line_metrics if m[4] > 0]
            if non_empty:
                max_h = max((m[4] for m in non_empty), default=new_size)
                spacing = int(max(0.0, options.line_spacing_scale) * max_h)
                total_h += spacing * (len(non_empty) - 1)
            y = y_top
            # header left, content right; truncate to avoid overlap
            region_left = lr_margin
            region_right = options.width - lr_margin
            total_w = max(0, region_right - region_left)
            min_gap = int(max(8, options.font_size * 0.40))
            def _elide_pil(s: str, max_w: int) -> tuple[str, int]:
                if max_w <= 0:
                    return ("", 0)
                w = draw.textlength(s, font=font)
                if w <= max_w:
                    return (s, int(w))
                ell = '…'
                ell_w = int(draw.textlength(ell, font=font))
                if ell_w > max_w:
                    return ("", 0)
                # binary search for fit
                lo, hi = 0, len(s)
                best = ("", 0)
                while lo <= hi:
                    mid = (lo + hi) // 2
                    cand = s[:mid] + ell
                    cw = int(draw.textlength(cand, font=font))
                    if cw <= max_w:
                        best = (cand, cw)
                        lo = mid + 1
                    else:
                        hi = mid - 1
                return best
            for lab, val, wL, wR, h in line_metrics:
                if h == 0:
                    continue
                if y + h > options.height - margin:
                    break
                max_val_w = max(0, total_w - min_gap)
                val_draw, wR2 = _elide_pil(val, max_val_w)
                avail_left = max(0, total_w - min_gap - wR2)
                lab_draw, wL2 = _elide_pil(lab, avail_left)
                xL = region_left
                xR = region_right - wR2
                draw.text((xL, y), lab_draw, font=font, fill=text_color)
                draw.text((xR, y), val_draw, font=font, fill=text_color)
                y += h + spacing
        else:
            # Fallback: anchor within a region below QR (legacy)
            qr_bottom = qr_y + qr_target_w
            region_left = lr_margin
            region_right = options.width - lr_margin
            region_top = qr_bottom + top_gap
            region_bottom = options.height - bottom_margin
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

        # dynamic divider based on x control point (or center)
        if options.text_point is not None:
            xn, _ = options.text_point
            divider = int(max(0.0, min(1.0, float(xn))) * options.width)
        else:
            divider = (region_left + region_right) // 2
        inner_pad = max(4, int(options.font_size * 0.20))
        for lab, val, wL, wR, h in line_metrics:
            if h == 0:
                continue
            # two-column positions around divider
            xL = divider - inner_pad - wL
            xR = divider + inner_pad
            xL = max(region_left, min(region_right - wL, xL))
            xR = max(region_left, min(region_right - wR, xR))
            if y + h > region_bottom:
                break
            draw.text((xL, y), lab, font=font, fill=text_color)
            draw.text((xR, y), val, font=font, fill=text_color)
            y += h + spacing

        fn = f"{a.id or a.name}.png".replace("/", "_").replace("\\", "_")
        canvas.save(out / fn, format="PNG")
        count += 1
    return count
