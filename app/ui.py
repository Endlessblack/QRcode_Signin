from __future__ import annotations

import sys
from dataclasses import dataclass
from collections import deque
from pathlib import Path
from typing import Optional

import cv2
import json
import qrcode
from PyQt6 import QtCore, QtGui, QtWidgets

from .config import AppConfig
from . import __version__
from .google_sheets import GoogleSheetsClient
from .logger import setup_logging
from .qr_tools import (
    Attendee,
    export_template_csv,
    generate_qr_images,
    load_attendees_csv,
    parse_qr_payload,
    DesignOptions,
    generate_qr_posters,
)


class WorkerAppendSheet(QtCore.QObject):
    finished = QtCore.pyqtSignal(dict)
    error = QtCore.pyqtSignal(str)

    def __init__(self, cfg: AppConfig, payload: dict):
        super().__init__()
        self.cfg = cfg
        self.payload = payload

    @QtCore.pyqtSlot()
    def run(self):
        import traceback
        log = setup_logging(self.cfg.debug)
        try:
            log.info("[API] Connecting to Google Sheet ...")
            client = GoogleSheetsClient(self.cfg.credentials_path, self.cfg.spreadsheet_id, self.cfg.worksheet_name)
            client.append_signin(self.payload)
            log.info("[API] Append success for id=%s name=%s", self.payload.get("id"), self.payload.get("name"))
            self.finished.emit(self.payload)
        except Exception:
            err = traceback.format_exc()
            log.error("[API] Append failed: %s", err)
            self.error.emit(err)


class InteractivePreview(QtWidgets.QLabel):
    regionChanged = QtCore.pyqtSignal(tuple)

    def __init__(self, parent: QtWidgets.QWidget | None = None):
        super().__init__(parent)
        self.setMouseTracking(True)
        self._canvas_w = 1080
        self._canvas_h = 1350
        self._norm_rect = [0.05, 0.60, 0.95, 0.95]  # x0,y0,x1,y1 normalized
        self._dragging = False
        self._resizing = False
        self._last_pos = QtCore.QPointF()

    def setCanvasSize(self, w: int, h: int):
        self._canvas_w = max(1, int(w))
        self._canvas_h = max(1, int(h))
        self.update()

    def setNormRect(self, rect: tuple[float, float, float, float]):
        x0, y0, x1, y1 = rect
        self._norm_rect = [float(x0), float(y0), float(x1), float(y1)]
        self._clamp_rect()
        self.update()

    def normRect(self) -> tuple[float, float, float, float]:
        return tuple(self._norm_rect)

    def _pixmap_rect(self) -> QtCore.QRectF:
        if self._canvas_w <= 0 or self._canvas_h <= 0:
            return QtCore.QRectF(0, 0, self.width(), self.height())
        scale = min(self.width() / self._canvas_w, self.height() / self._canvas_h)
        sw = self._canvas_w * scale
        sh = self._canvas_h * scale
        left = (self.width() - sw) / 2
        top = (self.height() - sh) / 2
        return QtCore.QRectF(left, top, sw, sh)

    def _label_to_norm(self, pos: QtCore.QPointF) -> QtCore.QPointF:
        r = self._pixmap_rect()
        if r.width() <= 0 or r.height() <= 0:
            return QtCore.QPointF(0, 0)
        x = (pos.x() - r.left()) / r.width()
        y = (pos.y() - r.top()) / r.height()
        x = max(0.0, min(1.0, x))
        y = max(0.0, min(1.0, y))
        return QtCore.QPointF(x, y)

    def _clamp_rect(self):
        x0, y0, x1, y1 = self._norm_rect
        x0 = max(0.0, min(1.0, x0))
        y0 = max(0.0, min(1.0, y0))
        x1 = max(0.0, min(1.0, x1))
        y1 = max(0.0, min(1.0, y1))
        if x1 < x0:
            x0, x1 = x1, x0
        if y1 < y0:
            y0, y1 = y1, y0
        # minimum size
        if x1 - x0 < 0.05:
            x1 = min(1.0, x0 + 0.05)
        if y1 - y0 < 0.05:
            y1 = min(1.0, y0 + 0.05)
        self._norm_rect = [x0, y0, x1, y1]

    def paintEvent(self, ev: QtGui.QPaintEvent):
        super().paintEvent(ev)
        # overlay red rectangle
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
        pen = QtGui.QPen(QtGui.QColor(255, 64, 64), 2)
        p.setPen(pen)
        r = self._pixmap_rect()
        x0, y0, x1, y1 = self._norm_rect
        rx0 = r.left() + r.width() * x0
        ry0 = r.top() + r.height() * y0
        rx1 = r.left() + r.width() * x1
        ry1 = r.top() + r.height() * y1
        p.drawRect(QtCore.QRectF(QtCore.QPointF(rx0, ry0), QtCore.QPointF(rx1, ry1)))

    def mousePressEvent(self, ev: QtGui.QMouseEvent):
        self._last_pos = ev.position()
        if ev.modifiers() & QtCore.Qt.KeyboardModifier.ShiftModifier:
            self._resizing = True
        else:
            self._dragging = True

    def mouseMoveEvent(self, ev: QtGui.QMouseEvent):
        if not (self._dragging or self._resizing):
            return
        cur = ev.position()
        prev = self._last_pos
        self._last_pos = cur
        prev_n = self._label_to_norm(prev)
        cur_n = self._label_to_norm(cur)
        dx = cur_n.x() - prev_n.x()
        dy = cur_n.y() - prev_n.y()
        x0, y0, x1, y1 = self._norm_rect
        if self._resizing:
            # resize from bottom-right by default
            x1 += dx
            y1 += dy
        else:
            x0 += dx; x1 += dx
            y0 += dy; y1 += dy
        self._norm_rect = [x0, y0, x1, y1]
        self._clamp_rect()
        self.update()
        self.regionChanged.emit(tuple(self._norm_rect))

    def mouseReleaseEvent(self, ev: QtGui.QMouseEvent):
        if self._dragging or self._resizing:
            self._dragging = False
            self._resizing = False
            self.regionChanged.emit(tuple(self._norm_rect))


class GenerateTab(QtWidgets.QWidget):
    def __init__(self, cfg: AppConfig):
        super().__init__()
        self.cfg = cfg
        self._build()

    def _build(self):
        layout = QtWidgets.QVBoxLayout(self)

        # Row: file + browse
        file_layout = QtWidgets.QHBoxLayout()
        self.file_edit = QtWidgets.QLineEdit()
        self.file_edit.setPlaceholderText("選擇名單 CSV 檔 (含 id,name)...")
        btn_browse = QtWidgets.QPushButton("選擇檔案")
        btn_browse.clicked.connect(self._choose_file)
        file_layout.addWidget(self.file_edit)
        file_layout.addWidget(btn_browse)
        layout.addLayout(file_layout)

        # Row: output folder
        out_layout = QtWidgets.QHBoxLayout()
        self.out_edit = QtWidgets.QLineEdit(self.cfg.qr_folder)
        btn_out = QtWidgets.QPushButton("輸出資料夾")
        btn_out.clicked.connect(self._choose_out)
        out_layout.addWidget(self.out_edit)
        out_layout.addWidget(btn_out)
        layout.addLayout(out_layout)

        # Row: event name
        event_layout = QtWidgets.QHBoxLayout()
        self.event_edit = QtWidgets.QLineEdit(self.cfg.event_name)
        self.event_edit.setPlaceholderText("活動名稱")
        event_layout.addWidget(QtWidgets.QLabel("活動名稱"))
        event_layout.addWidget(self.event_edit)
        layout.addLayout(event_layout)

        # Design panel
        design_group = QtWidgets.QGroupBox("圖面設計（預設 4:5 1080x1350）")
        grid = QtWidgets.QGridLayout()

        self.cb_use_design = QtWidgets.QCheckBox("使用設計版輸出（含文字/顏色/字型）")
        self.cb_use_design.setChecked(True)

        self.sp_width = QtWidgets.QSpinBox(); self.sp_width.setRange(300, 4000); self.sp_width.setValue(1080)
        self.sp_height = QtWidgets.QSpinBox(); self.sp_height.setRange(300, 4000); self.sp_height.setValue(1350)
        self.sl_qr_ratio = QtWidgets.QSlider(QtCore.Qt.Orientation.Horizontal)
        self.sl_qr_ratio.setRange(10, 100); self.sl_qr_ratio.setValue(70)
        self.lb_qr_ratio = QtWidgets.QLabel("70%")
        self.sl_qr_ratio.valueChanged.connect(lambda v: self.lb_qr_ratio.setText(f"{v}%"))

        self.ed_bg = QtWidgets.QLineEdit("#FFFFFF"); self.bg_input = self._attach_color_button(self.ed_bg)
        self.ed_qr = QtWidgets.QLineEdit("#000000"); self.qr_input = self._attach_color_button(self.ed_qr)
        self.ed_text = QtWidgets.QLineEdit("#000000"); self.text_input = self._attach_color_button(self.ed_text)

        # Font combo (installed fonts)
        self.font_combo = QtWidgets.QFontComboBox()
        self.lb_font_meta = QtWidgets.QLabel("")
        self.sp_font_size = QtWidgets.QSpinBox(); self.sp_font_size.setRange(10, 500); self.sp_font_size.setValue(48)
        
        # Optional background image to replace solid color
        self.ed_bg_img = QtWidgets.QLineEdit("")
        btn_bg_img = QtWidgets.QPushButton("選擇圖面…")
        btn_bg_img.clicked.connect(self._choose_bg_image)

        # Text block fine-tune
        self.cb_text_align = QtWidgets.QComboBox(); self.cb_text_align.addItems(["靠左", "置中", "靠右"]); self.cb_text_align.setCurrentText("置中")
        self.sp_text_margin = QtWidgets.QSpinBox(); self.sp_text_margin.setRange(0, 400); self.sp_text_margin.setValue(40)
        self.dsb_line_spacing = QtWidgets.QDoubleSpinBox(); self.dsb_line_spacing.setRange(0.0, 2.0); self.dsb_line_spacing.setSingleStep(0.1); self.dsb_line_spacing.setValue(0.4)
        self.cb_auto_fit = QtWidgets.QCheckBox("自動縮放文字以適配可用區域")
        self.cb_auto_fit.setChecked(True)

        r = 0
        grid.addWidget(self.cb_use_design, r, 0, 1, 3); r += 1
        grid.addWidget(QtWidgets.QLabel("寬度"), r, 0); grid.addWidget(self.sp_width, r, 1); r += 1
        grid.addWidget(QtWidgets.QLabel("高度"), r, 0); grid.addWidget(self.sp_height, r, 1); r += 1
        grid.addWidget(QtWidgets.QLabel("QR 寬度占比"), r, 0); grid.addWidget(self.sl_qr_ratio, r, 1); grid.addWidget(self.lb_qr_ratio, r, 2); r += 1
        grid.addWidget(QtWidgets.QLabel("背景色"), r, 0); grid.addWidget(self.bg_input, r, 1); r += 1
        grid.addWidget(QtWidgets.QLabel("QR 顏色"), r, 0); grid.addWidget(self.qr_input, r, 1); r += 1
        grid.addWidget(QtWidgets.QLabel("文字顏色"), r, 0); grid.addWidget(self.text_input, r, 1); r += 1
        grid.addWidget(QtWidgets.QLabel("字型"), r, 0); 
        hfont = QtWidgets.QHBoxLayout(); hfont.addWidget(self.font_combo)
        wfont = QtWidgets.QWidget(); wfont.setLayout(hfont)
        grid.addWidget(wfont, r, 1, 1, 2); r += 1
        grid.addWidget(QtWidgets.QLabel("字型大小"), r, 0); grid.addWidget(self.sp_font_size, r, 1); r += 1
        grid.addWidget(QtWidgets.QLabel("實際字型"), r, 0); grid.addWidget(self.lb_font_meta, r, 1, 1, 2); r += 1
        grid.addWidget(QtWidgets.QLabel("圖面路徑(可替代背景色)"), r, 0)
        hbg = QtWidgets.QHBoxLayout(); hbg.addWidget(self.ed_bg_img); hbg.addWidget(btn_bg_img)
        wbg = QtWidgets.QWidget(); wbg.setLayout(hbg)
        grid.addWidget(wbg, r, 1, 1, 2); r += 1

        grid.addWidget(QtWidgets.QLabel("文字對齊"), r, 0); grid.addWidget(self.cb_text_align, r, 1); r += 1
        grid.addWidget(QtWidgets.QLabel("文字邊界"), r, 0); grid.addWidget(self.sp_text_margin, r, 1); r += 1
        grid.addWidget(QtWidgets.QLabel("行距係數"), r, 0); grid.addWidget(self.dsb_line_spacing, r, 1); r += 1
        grid.addWidget(self.cb_auto_fit, r, 0, 1, 2); r += 1

        design_group.setLayout(grid)
        # Place design panel and preview side-by-side
        row = QtWidgets.QHBoxLayout()
        row.addWidget(design_group, 1)
        self.preview_label = InteractivePreview()
        self.preview_label.setMinimumSize(360, 360)
        self.preview_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setStyleSheet("background:#1c1c1c;border:1px solid #333;border-radius:6px")
        self.preview_label.regionChanged.connect(self._on_region_changed)
        row.addWidget(self.preview_label, 1)
        layout.addLayout(row)

        # Row: button (only generate here)
        btn_layout = QtWidgets.QHBoxLayout()
        self.btn_generate = QtWidgets.QPushButton("批次產生 QR Code")
        self.btn_generate.clicked.connect(self._generate)
        btn_layout.addStretch(1)
        btn_layout.addWidget(self.btn_generate)
        layout.addLayout(btn_layout)

        # Wire live preview updates
        self.sp_width.valueChanged.connect(self._preview)
        self.sp_height.valueChanged.connect(self._preview)
        self.sl_qr_ratio.valueChanged.connect(self._preview)
        self.ed_bg.textChanged.connect(self._preview)
        self.ed_qr.textChanged.connect(self._preview)
        self.ed_text.textChanged.connect(self._preview)
        self.font_combo.currentFontChanged.connect(lambda *_: self._preview())
        self.sp_font_size.valueChanged.connect(self._preview)
        self.ed_bg_img.textChanged.connect(self._preview)
        self.cb_text_align.currentIndexChanged.connect(self._preview)
        self.sp_text_margin.valueChanged.connect(self._preview)
        self.dsb_line_spacing.valueChanged.connect(self._preview)
        self.cb_auto_fit.toggled.connect(self._preview)

        # Status
        self.status = QtWidgets.QLabel()
        layout.addWidget(self.status)
        layout.addStretch(1)
        # Live preview triggers
        self.sp_width.valueChanged.connect(self._preview)
        self.sp_height.valueChanged.connect(self._preview)
        self.sl_qr_ratio.valueChanged.connect(self._preview)
        self.ed_bg.textChanged.connect(self._preview)
        self.ed_qr.textChanged.connect(self._preview)
        self.ed_text.textChanged.connect(self._preview)
        self.font_combo.currentFontChanged.connect(lambda *_: self._preview())
        self.sp_font_size.valueChanged.connect(self._preview)
        self.ed_bg_img.textChanged.connect(self._preview)
        self.cb_text_align.currentIndexChanged.connect(self._preview)
        self.sp_text_margin.valueChanged.connect(self._preview)
        self.dsb_line_spacing.valueChanged.connect(self._preview)
        self.cb_auto_fit.toggled.connect(self._preview)
        QtCore.QTimer.singleShot(0, self._preview)
        # initial preview
        QtCore.QTimer.singleShot(10, self._preview)

    def _choose_file(self):
        fn, _ = QtWidgets.QFileDialog.getOpenFileName(self, "選擇 CSV", str(Path.cwd()), "CSV (*.csv)")
        if fn:
            self.file_edit.setText(fn)

    def _choose_out(self):
        dn = QtWidgets.QFileDialog.getExistingDirectory(self, "選擇輸出資料夾", str(Path.cwd()))
        if dn:
            self.out_edit.setText(dn)

    # Export template moved to TemplateTab
    
    def _generate(self):
        csv_path = self.file_edit.text().strip()
        out_dir = self.out_edit.text().strip() or self.cfg.qr_folder
        event = self.event_edit.text().strip() or self.cfg.event_name
        try:
            attendees = load_attendees_csv(csv_path)
            if self.cb_use_design.isChecked():
                opts = DesignOptions(
                    width=int(self.sp_width.value()),
                    height=int(self.sp_height.value()),
                    qr_ratio=float(self.sl_qr_ratio.value()) / 100.0,
                    bg_color=self.ed_bg.text().strip(),
                    qr_color=self.ed_qr.text().strip(),
                    text_color=self.ed_text.text().strip(),
                    font_family=self.font_combo.currentFont().family(),
                    font_size=int(self.sp_font_size.value()),
                    bg_image_path=self.ed_bg_img.text().strip() or None,
                    text_align=self._map_align(),
                    text_margin=int(self.sp_text_margin.value()),
                    line_spacing_scale=float(self.dsb_line_spacing.value()),
                    auto_fit_text=bool(self.cb_auto_fit.isChecked()),
                )
                if hasattr(self, 'text_region_norm') and self.text_region_norm:
                    opts.text_region = self.text_region_norm
                count = generate_qr_posters(attendees, event, out_dir, opts)
            else:
                count = generate_qr_images(attendees, event, out_dir)
            self.status.setText(f"完成產生 {count} 張 圖片 → {out_dir}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "發生錯誤", str(e))

    def _attach_color_button(self, edit: QtWidgets.QLineEdit) -> QtWidgets.QWidget:
        btn = QtWidgets.QPushButton("…")
        btn.setFixedWidth(28)
        btn.clicked.connect(lambda: self._pick_color_into(edit))
        lay = QtWidgets.QHBoxLayout(); lay.setContentsMargins(0,0,0,0)
        lay.addWidget(edit); lay.addWidget(btn)
        w = QtWidgets.QWidget(); w.setLayout(lay)
        return w

    def _pick_color_into(self, edit: QtWidgets.QLineEdit):
        c = QtWidgets.QColorDialog.getColor(QtGui.QColor(edit.text().strip() or "#000000"), self, "選擇顏色")
        if c.isValid():
            edit.setText(c.name())

    def _choose_bg_image(self):
        fn, _ = QtWidgets.QFileDialog.getOpenFileName(self, "選擇圖面", str(Path.cwd()), "Image Files (*.png *.jpg *.jpeg)")
        if fn:
            self.ed_bg_img.setText(fn)

    def _preview(self):
        # Build a sample attendee from first row of CSV if present; else placeholders
        sample = Attendee(id="ID", name="NAME", extra={"salon": "SALON", "seller": "SELLER"})
        opts = DesignOptions(
            width=int(self.sp_width.value()),
            height=int(self.sp_height.value()),
            qr_ratio=float(self.sl_qr_ratio.value()) / 100.0,
            bg_color=self.ed_bg.text().strip(),
            qr_color=self.ed_qr.text().strip(),
            text_color=self.ed_text.text().strip(),
            font_family=self.font_combo.currentFont().family(),
            font_size=int(self.sp_font_size.value()),
            bg_image_path=self.ed_bg_img.text().strip() or None,
            text_align=self._map_align(),
            text_margin=int(self.sp_text_margin.value()),
            line_spacing_scale=float(self.dsb_line_spacing.value()),
            auto_fit_text=bool(self.cb_auto_fit.isChecked()),
        )
        if hasattr(self, 'text_region_norm') and self.text_region_norm:
            opts.text_region = self.text_region_norm
        # Render a single poster to memory
        try:
            # Reuse generator logic but not writing to disk: replicate steps
            payload = json.dumps({
                "id": sample.id,
                "name": sample.name,
                "event": self.event_edit.text().strip() or self.cfg.event_name,
                "extra": sample.extra,
            }, ensure_ascii=False)
            # QR
            qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M)
            qr.add_data(payload)
            qr.make(fit=True)
            qr_img = qr.make_image(fill_color=opts.qr_color, back_color="#FFFFFF").convert("RGB")

            from PIL import Image
            from PIL.ImageQt import ImageQt
            from PIL import ImageDraw as PILImageDraw, ImageFont as PILImageFont

            def _hex_to_rgb_local(hex_str: str):
                s = hex_str.strip()
                if s.startswith('#'):
                    s = s[1:]
                if len(s) == 3:
                    s = ''.join(ch*2 for ch in s)
                try:
                    return (int(s[0:2],16), int(s[2:4],16), int(s[4:6],16))
                except Exception:
                    return (0,0,0)

            # background: image or color
            if opts.bg_image_path and Path(opts.bg_image_path).exists():
                bg = Image.open(opts.bg_image_path).convert("RGB")
                sw, sh = bg.size
                tw, th = opts.width, opts.height
                scale = max(tw / sw, th / sh)
                bg = bg.resize((int(sw*scale), int(sh*scale)), Image.Resampling.LANCZOS)
                left = (bg.width - tw)//2; top = (bg.height - th)//2
                canvas = bg.crop((left, top, left+tw, top+th))
            else:
                canvas = Image.new("RGB", (opts.width, opts.height), _hex_to_rgb_local(opts.bg_color))
            draw = PILImageDraw.Draw(canvas)
            qr_target_w = int(opts.width * max(0.1, min(1.0, opts.qr_ratio)))
            qr_target_w = max(50, min(qr_target_w, opts.width - 80))
            qr_resized = qr_img.resize((qr_target_w, qr_target_w), Image.Resampling.LANCZOS)
            qr_x = (opts.width - qr_target_w) // 2
            qr_y = 40
            canvas.paste(qr_resized, (qr_x, qr_y))
            # font
            # Use the same font loader as generator and show resolved info
            try:
                from app.qr_tools import get_font_with_meta as get_font_meta
                font, meta = get_font_meta(opts.font_size, opts.font_family, None)
                try:
                    from pathlib import Path as _P
                    self.lb_font_meta.setText(f"{meta.get('name','')} ({_P(meta.get('path','')).name}#{meta.get('index',0)})")
                except Exception:
                    self.lb_font_meta.setText(str(meta))
            except Exception:
                from PIL import ImageFont as PILImageFont2
                font = PILImageFont2.load_default()
                self.lb_font_meta.setText("default")
            text_color = _hex_to_rgb_local(opts.text_color)
            lines = [
                f"ID: {sample.id}",
                f"姓名: {sample.name}",
                f"Salon: {sample.extra.get('salon','')}",
                f"Seller: {sample.extra.get('seller','')}",
            ]
            # Place text block anchored in area below QR only
            # spacing based on measured max line height rather than requested size
            metrics = []
            total_h = 0
            for line in lines:
                bbox = draw.textbbox((0,0), line, font=font)
                w = bbox[2]-bbox[0]; h=bbox[3]-bbox[1]
                metrics.append((line, w, h))
                total_h += h
            max_h = max((h for _,_,h in metrics if h>0), default=opts.font_size)
            spacing = int(max(0.0, opts.line_spacing_scale) * max_h)
            total_h += spacing * (len([m for m in metrics if m[2]>0]) - 1)
            margin = max(0, int(opts.text_margin))
            if hasattr(self, 'text_region_norm') and self.text_region_norm:
                x0n, y0n, x1n, y1n = self.text_region_norm
                region_top = int(y0n * opts.height)
                region_bottom = int(y1n * opts.height)
                region_left = int(x0n * opts.width)
                region_right = int(x1n * opts.width)
            else:
                qr_bottom = qr_y + qr_target_w
                region_top = qr_bottom + margin
                region_bottom = opts.height - margin
                region_left = margin
                region_right = opts.width - margin
            region_height = max(0, region_bottom - region_top)
            # default anchor to top within computed region
            y = region_top
            for line, w, h in metrics:
                if h == 0:
                    continue
                if opts.text_align == 'left':
                    x = region_left
                elif opts.text_align == 'right':
                    x = max(region_left, region_right - w)
                else:
                    x = (region_left + region_right - w)//2
                if y + h > region_bottom:
                    break
                draw.text((x,y), line, fill=text_color, font=font)
                y += h + spacing

            qim = ImageQt(canvas)
            pix = QtGui.QPixmap.fromImage(qim).scaled(self.preview_label.width(), self.preview_label.height(), QtCore.Qt.AspectRatioMode.KeepAspectRatio)
            self.preview_label.setPixmap(pix)
            self.preview_label.setCanvasSize(opts.width, opts.height)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "預覽失敗", str(e))

    def _on_region_changed(self, norm_rect: tuple):
        self.text_region_norm = tuple(norm_rect)
        self._preview()

    def _map_align(self) -> str:
        t = self.cb_text_align.currentText()
        return {'靠左':'left','置中':'center','靠右':'right'}.get(t, 'center')


class ScanTab(QtWidgets.QWidget):
    def __init__(self, cfg: AppConfig, sheets: GoogleSheetsClient):
        super().__init__()
        self.cfg = cfg
        self.sheets = sheets
        self.cap: Optional[cv2.VideoCapture] = None
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self._next_frame)
        self.detector = cv2.QRCodeDetector()
        # Queue-based processing: scan enqueues; worker consumes
        self.queue = deque()
        self._api_processing = False
        self._cooldown = False  # short debounce to avoid duplicate frames
        self._failsafe_timer: Optional[QtCore.QTimer] = None
        self._jobs: list[tuple[QtCore.QThread, WorkerAppendSheet]] = []
        self._build()
        self.log = setup_logging(self.cfg.debug)

    def _build(self):
        layout = QtWidgets.QVBoxLayout(self)

        # Camera preview
        self.preview = QtWidgets.QLabel()
        self.preview.setFixedHeight(360)
        self.preview.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.preview.setStyleSheet("background:#222;color:#bbb")
        layout.addWidget(self.preview)

        # Controls
        ctrl = QtWidgets.QHBoxLayout()
        self.btn_start = QtWidgets.QPushButton("啟動相機")
        self.btn_stop = QtWidgets.QPushButton("停止")
        self.btn_start.clicked.connect(self.start_camera)
        self.btn_stop.clicked.connect(self.stop_camera)
        ctrl.addWidget(self.btn_start)
        ctrl.addWidget(self.btn_stop)
        ctrl.addStretch(1)
        layout.addLayout(ctrl)

        # Live log panel
        self.log_view = QtWidgets.QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setFixedHeight(120)
        self.log_view.setStyleSheet("QTextEdit { background:#111;color:#9cdcfe;border:1px solid #333;border-radius:6px; }")
        layout.addWidget(self.log_view)

        # Table of scans
        self.table = QtWidgets.QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["時間", "ID", "姓名", "內容"])
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)

    def start_camera(self):
        if self.cap is not None:
            return
        idx = int(self.cfg.camera_index)
        self.cap = cv2.VideoCapture(idx, cv2.CAP_DSHOW)
        if not self.cap.isOpened():
            self.cap.release()
            self.cap = None
            msg = f"相機無法開啟，索引：{idx}"
            self._log(msg)
            QtWidgets.QMessageBox.critical(self, "相機無法開啟", msg)
            return
        self._log(f"相機已啟動（索引 {idx}）")
        self.timer.start(30)

    def stop_camera(self):
        self.timer.stop()
        if self.cap is not None:
            self.cap.release()
            self.cap = None
        self.preview.clear()
        self._log("相機已停止")

    def _next_frame(self):
        if self.cap is None:
            return
        ok, frame = self.cap.read()
        if not ok:
            return
        # Show preview
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w, ch = rgb.shape
        img = QtGui.QImage(rgb.data, w, h, ch * w, QtGui.QImage.Format.Format_RGB888)
        pix = QtGui.QPixmap.fromImage(img).scaled(self.preview.width(), self.preview.height(), QtCore.Qt.AspectRatioMode.KeepAspectRatio)
        self.preview.setPixmap(pix)

        # Detect QR with short cooldown debounce
        if self._cooldown:
            return
        data, points, _ = self.detector.detectAndDecode(frame)
        if data:
            # start short cooldown (does not block API queue)
            self._cooldown = True
            QtCore.QTimer.singleShot(800, lambda: setattr(self, "_cooldown", False))
            self._handle_qr_text(data)

    def _handle_qr_text(self, data: str):
        payload = parse_qr_payload(data, self.cfg.event_name)
        self._log(f"偵測到 QR：{str(payload)[:120]}")

        # Append UI row
        from datetime import datetime

        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(ts))
        self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(str(payload.get("id", ""))))
        self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(str(payload.get("name", ""))))
        self.table.setItem(row, 3, QtWidgets.QTableWidgetItem(data))

        # Enqueue payload and process
        self.queue.append(payload)
        self._log(f"已加入佇列，待寫入（佇列長度 {len(self.queue)}）")
        self._start_next_job()

    def _start_next_job(self):
        if self._api_processing or not self.queue:
            return
        payload = self.queue.popleft()
        self._api_processing = True

        thread = QtCore.QThread()
        worker = WorkerAppendSheet(self.cfg, payload)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        worker.error.connect(self._on_error)
        worker.finished.connect(lambda _p=payload: self._on_api_success(_p))
        thread.finished.connect(self._on_done)
        # Keep strong refs until finished
        self._jobs.append((thread, worker))
        thread.finished.connect(lambda: self._jobs.remove((thread, worker)) if (thread, worker) in self._jobs else None)
        self._log("[API] 開始寫入 Google Sheet ...")
        thread.start()

        # Failsafe per job: reset processing if API hangs
        if self._failsafe_timer is not None:
            try:
                self._failsafe_timer.stop()
            except Exception:
                pass
        self._failsafe_timer = QtCore.QTimer(self)
        self._failsafe_timer.setSingleShot(True)
        self._failsafe_timer.timeout.connect(self._job_timeout)
        self._failsafe_timer.start(10000)  # 10s

    def _on_error(self, msg: str):
        self._log("[API] 失敗：" + msg)
        QtWidgets.QMessageBox.warning(self, "寫入 Google 失敗", msg)
        self._on_done()

    def _on_done(self):
        QtWidgets.QApplication.beep()
        # cancel failsafe if it’s pending
        if self._failsafe_timer is not None:
            try:
                self._failsafe_timer.stop()
            except Exception:
                pass
            self._failsafe_timer = None
        # mark API slot available and process next
        self._api_processing = False
        self._start_next_job()

    def _on_api_success(self, payload: dict):
        self._log(f"[API] 成功寫入：id={payload.get('id','')}, name={payload.get('name','')}")

    def _log(self, text: str):
        from datetime import datetime
        ts = datetime.now().strftime("%H:%M:%S")
        self.log.info(text)
        self.log_view.append(f"[{ts}] {text}")

    def _job_timeout(self):
        self._log("[API] 寫入逾時，跳過此筆並繼續下一筆（10s）")
        self._api_processing = False
        self._start_next_job()


class SettingsTab(QtWidgets.QWidget):
    config_changed = QtCore.pyqtSignal()

    def __init__(self, cfg: AppConfig):
        super().__init__()
        self.cfg = cfg
        self._build()

    def _build(self):
        form = QtWidgets.QFormLayout(self)

        self.ed_credentials = QtWidgets.QLineEdit(self.cfg.credentials_path)
        btn_cred = QtWidgets.QPushButton("瀏覽...")
        btn_cred.clicked.connect(self._choose_credentials)
        h1 = QtWidgets.QHBoxLayout()
        h1.addWidget(self.ed_credentials)
        h1.addWidget(btn_cred)

        # Spreadsheet URL instead of ID; display mapped URL
        self.ed_spreadsheet = QtWidgets.QLineEdit(self._to_sheet_url(self.cfg.spreadsheet_id))
        self.ed_worksheet = QtWidgets.QLineEdit(self.cfg.worksheet_name)
        self.ed_event = QtWidgets.QLineEdit(self.cfg.event_name)
        # Camera combobox + refresh button
        self.cb_camera = QtWidgets.QComboBox()
        btn_cam_refresh = QtWidgets.QPushButton("刷新")
        btn_cam_refresh.clicked.connect(self._populate_cameras)
        cam_row = QtWidgets.QHBoxLayout()
        cam_row.addWidget(self.cb_camera)
        cam_row.addWidget(btn_cam_refresh)
        self._populate_cameras()

        form.addRow("憑證檔案", self._wrap(h1))
        form.addRow("試算表 URL", self.ed_spreadsheet)
        # Show parsed spreadsheet ID for clarity
        self.lb_sheet_id = QtWidgets.QLabel("")
        self.lb_sheet_id.setStyleSheet("color:#8ab4f8")
        self._update_sheet_id_label()
        self.ed_spreadsheet.textChanged.connect(self._update_sheet_id_label)
        form.addRow("解析出 ID", self.lb_sheet_id)
        form.addRow("工作表名稱", self.ed_worksheet)
        form.addRow("活動名稱", self.ed_event)
        form.addRow("相機來源", self._wrap(cam_row))

        # Debug toggle
        self.cb_debug = QtWidgets.QCheckBox("啟用除錯紀錄 (logs/app.log)")
        self.cb_debug.setChecked(self.cfg.debug)
        form.addRow(self.cb_debug)

        btn_save = QtWidgets.QPushButton("儲存設定")
        btn_test = QtWidgets.QPushButton("測試連線")
        btn_save.clicked.connect(self._save)
        btn_test.clicked.connect(self._test_connection)
        hb = QtWidgets.QHBoxLayout()
        hb.addWidget(btn_save)
        hb.addWidget(btn_test)
        hb.addStretch(1)
        form.addRow(self._wrap(hb))

    def _wrap(self, layout: QtWidgets.QLayout) -> QtWidgets.QWidget:
        w = QtWidgets.QWidget()
        w.setLayout(layout)
        return w

    def _choose_credentials(self):
        fn, _ = QtWidgets.QFileDialog.getOpenFileName(self, "選擇 credentials.json", str(Path.cwd()))
        if fn:
            self.ed_credentials.setText(fn)

    def _save(self):
        self.cfg.credentials_path = self.ed_credentials.text().strip()
        # Accept URL or ID, map to ID
        self.cfg.spreadsheet_id = self._extract_spreadsheet_id(self.ed_spreadsheet.text().strip())
        self.cfg.worksheet_name = self.ed_worksheet.text().strip()
        self.cfg.event_name = self.ed_event.text().strip()
        data = self.cb_camera.currentData()
        try:
            self.cfg.camera_index = int(data if data is not None else 0)
        except Exception:
            self.cfg.camera_index = 0
        self.cfg.debug = bool(self.cb_debug.isChecked())
        self.cfg.save()
        QtWidgets.QMessageBox.information(self, "設定", "已儲存")
        self.config_changed.emit()

    def _extract_spreadsheet_id(self, s: str) -> str:
        # Accept raw ID or full URL like https://docs.google.com/spreadsheets/d/<ID>/edit
        import re
        s = s.strip()
        if not s:
            return ""
        m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", s)
        if m:
            return m.group(1)
        return s  # assume already an ID

    def _to_sheet_url(self, sid: str) -> str:
        sid = (sid or "").strip()
        if not sid:
            return ""
        return f"https://docs.google.com/spreadsheets/d/{sid}/edit"

    def _update_sheet_id_label(self):
        sid = self._extract_spreadsheet_id(self.ed_spreadsheet.text())
        self.lb_sheet_id.setText(sid or "(未解析到 ID)")

    def _get_service_account_email(self) -> str:
        import json
        try:
            path = self.ed_credentials.text().strip() or self.cfg.credentials_path
            with open(path, "r", encoding="utf-8") as f:
                return str(json.load(f).get("client_email", ""))
        except Exception:
            return ""

    def _test_connection(self):
        sid = self._extract_spreadsheet_id(self.ed_spreadsheet.text().strip())
        email = self._get_service_account_email()
        from .google_sheets import GoogleSheetsClient
        client = GoogleSheetsClient(self.ed_credentials.text().strip() or self.cfg.credentials_path,
                                    sid,
                                    self.ed_worksheet.text().strip() or self.cfg.worksheet_name)
        try:
            client.connect()
        except Exception as e:
            msg = (
                f"連線失敗：{e}\n\n請確認：\n"
                f"1) 試算表已分享給服務帳戶 email：{email or '(無法讀取)'}（可編輯）\n"
                f"2) 試算表 URL/ID 正確（上方解析出 ID：{sid or '(無)'}）\n"
                f"3) 已在 GCP 專案啟用 Google Sheets API 與 Drive API"
            )
            QtWidgets.QMessageBox.critical(self, "測試連線失敗", msg)
        else:
            QtWidgets.QMessageBox.information(self, "測試連線成功", f"成功連線至試算表（ID={sid}）。\n服務帳戶：{email or '(未知)'}")

    def _get_camera_names(self) -> list[str] | None:
        try:
            from pygrabber.dshow_graph import FilterGraph  # type: ignore
            graph = FilterGraph()
            return list(graph.get_input_devices())
        except Exception:
            return None

    def _populate_cameras(self):
        sel_idx = int(self.cfg.camera_index)
        names = self._get_camera_names()
        found: list[tuple[int, str]] = []
        if names:
            for i, name in enumerate(names):
                cap = cv2.VideoCapture(i, cv2.CAP_DSHOW)
                if cap is not None and cap.isOpened():
                    found.append((i, name))
                    cap.release()
        else:
            for i in range(10):
                cap = cv2.VideoCapture(i, cv2.CAP_DSHOW)
                if cap is not None and cap.isOpened():
                    found.append((i, f"Camera {i}"))
                    cap.release()
        self.cb_camera.clear()
        if found:
            idx_list = []
            for idx, name in found:
                self.cb_camera.addItem(name, idx)
                idx_list.append(idx)
            if sel_idx in idx_list:
                self.cb_camera.setCurrentIndex(idx_list.index(sel_idx))
            else:
                self.cb_camera.setCurrentIndex(0)
        else:
            self.cb_camera.addItem("無可用相機 (預設 0)", 0)

class TemplateTab(QtWidgets.QWidget):
    def __init__(self, cfg: AppConfig):
        super().__init__()
        self.cfg = cfg
        self._build()

    def _build(self):
        layout = QtWidgets.QVBoxLayout(self)

        # Editor for custom fields
        group = QtWidgets.QGroupBox("自訂欄位（除了 id, name）")
        v = QtWidgets.QVBoxLayout()
        self.list_fields = QtWidgets.QListWidget()
        self.list_fields.addItems(self.cfg.extra_fields)
        self.list_fields.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        # Show black text as requested
        self.list_fields.setStyleSheet("QListWidget { background: #ffffff; color: #000000; }")
        v.addWidget(self.list_fields)

        h = QtWidgets.QHBoxLayout()
        btn_add = QtWidgets.QPushButton("新增")
        btn_edit = QtWidgets.QPushButton("編輯")
        btn_del = QtWidgets.QPushButton("刪除")
        btn_up = QtWidgets.QPushButton("上移")
        btn_down = QtWidgets.QPushButton("下移")
        h.addWidget(btn_add)
        h.addWidget(btn_edit)
        h.addWidget(btn_del)
        h.addStretch(1)
        h.addWidget(btn_up)
        h.addWidget(btn_down)
        v.addLayout(h)

        group.setLayout(v)
        layout.addWidget(group)

        # Buttons: Save fields and Export template
        btns = QtWidgets.QHBoxLayout()
        btn_save = QtWidgets.QPushButton("儲存欄位")
        btn_export = QtWidgets.QPushButton("輸出範本 CSV")
        btns.addWidget(btn_save)
        btns.addStretch(1)
        btns.addWidget(btn_export)
        layout.addLayout(btns)

        self.status = QtWidgets.QLabel()
        layout.addWidget(self.status)
        layout.addStretch(1)

        # Wire actions
        btn_add.clicked.connect(self._field_add)
        btn_edit.clicked.connect(self._field_edit)
        btn_del.clicked.connect(self._field_delete)
        btn_up.clicked.connect(lambda: self._field_move(-1))
        btn_down.clicked.connect(lambda: self._field_move(1))
        btn_save.clicked.connect(self._save_fields)
        btn_export.clicked.connect(self._export_template)

    def _save_fields(self):
        items = [self.list_fields.item(i).text() for i in range(self.list_fields.count())]
        self.cfg.extra_fields = items
        self.cfg.save()
        self.status.setText("已儲存自訂欄位到設定檔")

    def _export_template(self):
        fn, _ = QtWidgets.QFileDialog.getSaveFileName(self, "儲存範本", str(Path.cwd() / "attendees_template.csv"), "CSV (*.csv)")
        if fn:
            fields = [self.list_fields.item(i).text() for i in range(self.list_fields.count())]
            export_template_csv(fn, fields)
            self.status.setText(f"已輸出範本：{fn}")

    def _field_add(self):
        text, ok = QtWidgets.QInputDialog.getText(self, "新增欄位", "欄位名稱：")
        if ok:
            name = text.strip()
            if name and name not in ("id", "name"):
                existing = [self.list_fields.item(i).text() for i in range(self.list_fields.count())]
                if name not in existing:
                    self.list_fields.addItem(name)

    def _field_edit(self):
        it = self.list_fields.currentItem()
        if not it:
            return
        text, ok = QtWidgets.QInputDialog.getText(self, "編輯欄位", "欄位名稱：", text=it.text())
        if ok:
            name = text.strip()
            if name and name not in ("id", "name"):
                it.setText(name)

    def _field_delete(self):
        row = self.list_fields.currentRow()
        if row >= 0:
            self.list_fields.takeItem(row)

    def _field_move(self, delta: int):
        row = self.list_fields.currentRow()
        if row < 0:
            return
        new_row = row + delta
        if 0 <= new_row < self.list_fields.count():
            it = self.list_fields.takeItem(row)
            self.list_fields.insertItem(new_row, it)
            self.list_fields.setCurrentRow(new_row)


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, cfg: AppConfig):
        super().__init__()
        self.cfg = cfg
        self.setWindowTitle(f"QR 簽到 v{__version__}")
        self.resize(1000, 700)
        self._init_ui()

    def _init_ui(self):
        tabs = QtWidgets.QTabWidget()
        self.setCentralWidget(tabs)

        # Sheets client is created lazily to avoid blocking UI on startup
        self.sheets_client = GoogleSheetsClient(
            self.cfg.credentials_path, self.cfg.spreadsheet_id, self.cfg.worksheet_name
        )

        # Put Scan tab first as Home
        self.tab_scan = ScanTab(self.cfg, self.sheets_client)
        self.tab_generate = GenerateTab(self.cfg)
        self.tab_settings = SettingsTab(self.cfg)
        self.tab_template = TemplateTab(self.cfg)
        self.tab_settings.config_changed.connect(self._on_config_changed)

        tabs.addTab(self.tab_scan, "首頁簽到")
        tabs.addTab(self.tab_generate, "產生 QR Code")
        tabs.addTab(self.tab_template, "建立範本")
        tabs.addTab(self.tab_settings, "設定")

        # Style
        self._apply_style()

    def _apply_style(self):
        self.setStyleSheet(
            """
            QMainWindow { background: #111; }
            QWidget { color: #eee; font-size: 14px; }
            QLineEdit, QSpinBox, QTextEdit, QComboBox, QTableWidget, QTableView {
                background: #1c1c1c; color: #eee; border: 1px solid #333; border-radius: 6px; padding: 4px 6px;
            }
            QPushButton { background: #2d6cdf; color: white; border: none; padding: 8px 14px; border-radius: 8px; }
            QPushButton:hover { background: #3b78e7; }
            QPushButton:disabled { background: #555; }
            QTabWidget::pane { border: 1px solid #333; }
            /* Tab labels: unselected on white with black text */
            QTabBar::tab { background: #ffffff; color: #000000; padding: 8px 12px; border-top-left-radius: 6px; border-top-right-radius: 6px; }
            QTabBar::tab:selected { background: #2d6cdf; color: #ffffff; }
            QTabBar::tab:hover:!selected { background: #f2f2f2; }
            QHeaderView::section { background: #222; color: #aaa; padding: 6px; border: none; }
            QLabel { color: #bbb; }
            """
        )

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        # Persist template fields on exit so they’re available next launch
        try:
            if hasattr(self, "tab_template") and isinstance(self.tab_template, TemplateTab):
                items = [self.tab_template.list_fields.item(i).text() for i in range(self.tab_template.list_fields.count())]
                self.cfg.extra_fields = items
                self.cfg.save()
        finally:
            super().closeEvent(event)

    def _on_config_changed(self):
        # Recreate sheets client with new config
        self.sheets_client = GoogleSheetsClient(
            self.cfg.credentials_path, self.cfg.spreadsheet_id, self.cfg.worksheet_name
        )
