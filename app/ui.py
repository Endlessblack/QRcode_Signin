from __future__ import annotations

import sys
from dataclasses import dataclass
from collections import deque
from pathlib import Path
from typing import Optional

import cv2
from PyQt6 import QtCore, QtGui, QtWidgets

from .config import AppConfig
from .google_sheets import GoogleSheetsClient
from .logger import setup_logging
from .qr_tools import (
    Attendee,
    export_template_csv,
    generate_qr_images,
    load_attendees_csv,
    parse_qr_payload,
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

        # Row: button (only generate here)
        btn_layout = QtWidgets.QHBoxLayout()
        self.btn_generate = QtWidgets.QPushButton("批次產生 QR Code")
        self.btn_generate.clicked.connect(self._generate)
        btn_layout.addStretch(1)
        btn_layout.addWidget(self.btn_generate)
        layout.addLayout(btn_layout)

        # Status
        self.status = QtWidgets.QLabel()
        layout.addWidget(self.status)
        layout.addStretch(1)

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
            count = generate_qr_images(attendees, event, out_dir)
            self.status.setText(f"完成產生 {count} 張 QR Code 圖片 → {out_dir}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "發生錯誤", str(e))


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
        self.setWindowTitle("QR 簽到")
        self.resize(1000, 700)
        self._init_ui()

    def _init_ui(self):
        tabs = QtWidgets.QTabWidget()
        self.setCentralWidget(tabs)

        # Sheets client is created lazily to avoid blocking UI on startup
        self.sheets_client = GoogleSheetsClient(
            self.cfg.credentials_path, self.cfg.spreadsheet_id, self.cfg.worksheet_name
        )

        self.tab_generate = GenerateTab(self.cfg)
        self.tab_scan = ScanTab(self.cfg, self.sheets_client)
        self.tab_settings = SettingsTab(self.cfg)
        self.tab_template = TemplateTab(self.cfg)
        self.tab_settings.config_changed.connect(self._on_config_changed)

        tabs.addTab(self.tab_generate, "產生 QR Code")
        tabs.addTab(self.tab_template, "建立範本")
        tabs.addTab(self.tab_scan, "掃描簽到")
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
