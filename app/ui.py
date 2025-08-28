from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import cv2
from PyQt6 import QtCore, QtGui, QtWidgets

from .config import AppConfig
from .google_sheets import GoogleSheetsClient
from .qr_tools import (
    Attendee,
    export_template_csv,
    generate_qr_images,
    load_attendees_csv,
    parse_qr_payload,
)


class WorkerAppendSheet(QtCore.QObject):
    finished = QtCore.pyqtSignal()
    error = QtCore.pyqtSignal(str)

    def __init__(self, client: GoogleSheetsClient, payload: dict):
        super().__init__()
        self.client = client
        self.payload = payload

    @QtCore.pyqtSlot()
    def run(self):
        try:
            self.client.append_signin(self.payload)
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))


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

        # Row: buttons
        btn_layout = QtWidgets.QHBoxLayout()
        self.btn_template = QtWidgets.QPushButton("輸出範本")
        self.btn_template.clicked.connect(self._export_template)
        self.btn_generate = QtWidgets.QPushButton("批次產生 QR Code")
        self.btn_generate.clicked.connect(self._generate)
        btn_layout.addWidget(self.btn_template)
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

    def _export_template(self):
        fn, _ = QtWidgets.QFileDialog.getSaveFileName(self, "儲存範本", str(Path.cwd() / "attendees_template.csv"), "CSV (*.csv)")
        if fn:
            export_template_csv(fn)
            self.status.setText(f"已輸出範本：{fn}")

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
        self._busy = False
        self._build()

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
            QtWidgets.QMessageBox.critical(self, "相機無法開啟", f"請檢查相機索引：{idx}")
            return
        self.timer.start(30)

    def stop_camera(self):
        self.timer.stop()
        if self.cap is not None:
            self.cap.release()
            self.cap = None
        self.preview.clear()

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

        # Detect QR
        if self._busy:
            return
        data, points, _ = self.detector.detectAndDecode(frame)
        if data:
            self._busy = True
            self._handle_qr_text(data)

    def _handle_qr_text(self, data: str):
        payload = parse_qr_payload(data, self.cfg.event_name)

        # Append UI row
        from datetime import datetime

        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(ts))
        self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(str(payload.get("id", ""))))
        self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(str(payload.get("name", ""))))
        self.table.setItem(row, 3, QtWidgets.QTableWidgetItem(data))

        # Write to Google in background
        self.thread = QtCore.QThread()
        worker = WorkerAppendSheet(self.sheets, payload)
        worker.moveToThread(self.thread)
        self.thread.started.connect(worker.run)
        worker.finished.connect(self.thread.quit)
        worker.finished.connect(worker.deleteLater)
        worker.error.connect(self._on_error)
        self.thread.finished.connect(self._on_done)
        self.thread.start()

    def _on_error(self, msg: str):
        QtWidgets.QMessageBox.warning(self, "寫入 Google 失敗", msg)
        self._on_done()

    def _on_done(self):
        QtWidgets.QApplication.beep()
        # small cooldown to avoid multiple fires on same frame
        QtCore.QTimer.singleShot(800, self._reset_busy)

    def _reset_busy(self):
        self._busy = False


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

        self.ed_spreadsheet = QtWidgets.QLineEdit(self.cfg.spreadsheet_id)
        self.ed_worksheet = QtWidgets.QLineEdit(self.cfg.worksheet_name)
        self.ed_event = QtWidgets.QLineEdit(self.cfg.event_name)
        self.ed_camera = QtWidgets.QSpinBox()
        self.ed_camera.setRange(0, 10)
        self.ed_camera.setValue(int(self.cfg.camera_index))

        form.addRow("憑證檔案", self._wrap(h1))
        form.addRow("試算表 ID", self.ed_spreadsheet)
        form.addRow("工作表名稱", self.ed_worksheet)
        form.addRow("活動名稱", self.ed_event)
        form.addRow("相機索引", self.ed_camera)

        btn_save = QtWidgets.QPushButton("儲存設定")
        btn_save.clicked.connect(self._save)
        form.addRow(btn_save)

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
        self.cfg.spreadsheet_id = self.ed_spreadsheet.text().strip()
        self.cfg.worksheet_name = self.ed_worksheet.text().strip()
        self.cfg.event_name = self.ed_event.text().strip()
        self.cfg.camera_index = int(self.ed_camera.value())
        self.cfg.save()
        QtWidgets.QMessageBox.information(self, "設定", "已儲存")
        self.config_changed.emit()


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
        self.tab_settings.config_changed.connect(self._on_config_changed)

        tabs.addTab(self.tab_generate, "產生 QR Code")
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
            QHeaderView::section { background: #222; color: #aaa; padding: 6px; border: none; }
            QLabel { color: #bbb; }
            """
        )

    def _on_config_changed(self):
        # Recreate sheets client with new config
        self.sheets_client = GoogleSheetsClient(
            self.cfg.credentials_path, self.cfg.spreadsheet_id, self.cfg.worksheet_name
        )

