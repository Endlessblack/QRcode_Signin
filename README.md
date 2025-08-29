# Windows QR Code 簽到程式（Google Sheet 後端）

版本：v1.0.0（穩定版）

一個可在 Windows 執行的桌面簽到工具：
- 以本機名單（CSV）批次產生 QR Code 圖檔
- 使用攝影機掃描 QR Code，將簽到紀錄寫入 Google Sheet
- 介面簡潔，安裝步驟清楚

## 功能總覽
- 產生 QR Code：從 CSV 名單匯入，批次輸出 PNG 圖片。
- 掃描簽到：啟動攝影機，擷取 QR 內容，寫入雲端 Google Sheet。
- 設定頁：設定 Google 憑證、試算表 ID、工作表名稱、活動名稱與相機來源。
- 名單範本：程式可一鍵輸出範本 `attendees_template.csv`。

## 系統需求
- Windows 10/11（x64）
- Python 3.10+（建議 3.10 或 3.11）
- 攝影機（USB 或筆電內建）

## 安裝步驟
1) 安裝相依套件

```bash
python -m venv .venv
.\.venv\Scripts\pip install --upgrade pip
.\.venv\Scripts\pip install -r requirements.txt
```

2) 建立 Google Service Account（服務帳戶）
- 到 Google Cloud Console 建立專案並啟用 Google Sheets API、Google Drive API。
- 建立 Service Account，並建立金鑰（JSON）。
- 下載 JSON 憑證檔，存成 `credentials.json` 放在專案根目錄（或你指定的位置）。
- 到你的目標 Google 試算表，將該試算表分享給 Service Account 的 email（例如 `my-bot@project.iam.gserviceaccount.com`），給「可編輯」權限。

3) 建立設定與目錄結構（為單檔打包準備）
- 設定檔固定路徑：`./setting/config.json`
- OAuth 與 token 固定資料夾：`./client`
- 建議建立以下結構：

```
./setting/config.json
./client/client.json          (若使用 OAuth，檔名已固定)
./client/token.json           (程式首次登入後自動產生/更新)
```

- `config.json` 內容需包含：
  - `google.credentials_path`：服務帳戶 JSON 路徑
  - `google.spreadsheet_id`：試算表 ID（URL 中的那串長字串）
  - `google.worksheet_name`：工作表名稱（不存在會自動建立）
  - `event.name`：活動名稱（會寫入簽到紀錄）

4) 執行程式

```bash
.\.venv\Scripts\python -m app.main
```

## CSV 名單格式
- 以 UTF-8 編碼 CSV，至少需包含欄位：`id,name`
- 可加上 `email`、`company` 等自訂欄位，系統會包含在 QR 內容中。

範本輸出：
- 於「產生 QR Code」分頁點選「輸出範本」，將產生 `attendees_template.csv`

## QR 內容格式
- 程式使用 JSON 格式放入 QR：

```json
{
  "id": "A001",
  "name": "王小明",
  "event": "Your Event",
  "extra": { "email": "...", "company": "..." }
}
```

掃描時會解析 JSON；若非 JSON 也會以原字串寫入（`raw` 欄位）。

## Google Sheet 欄位（預設）
- `timestamp`：ISO 時間（本地時區）
- `event`：活動名稱
- `id`、`name`：從 QR 內容取得
- `raw`：若 QR 不是 JSON，原始字串放這
- 其餘欄位：會展開 `extra` 中的鍵值（如 `email`, `company`）

首次寫入會自動建立標題列；不存在的欄位會自動擴充。

## 常見問題
- 無法連到 Google：
  - 檢查 `credentials.json` 路徑與檔案內容
  - 試算表是否分享給服務帳戶 email
  - `spreadsheet_id` 是否正確
- 相機打不開：
  - 在設定頁調整相機索引（0,1,2...）
  - 檢查其他軟體是否佔用相機
- 無法安裝 OpenCV：
  - 先升級 pip，或使用 `python -m pip install --only-binary=:all: opencv-python`

## 專案結構
- `app/main.py`：程式進入點
- `app/ui.py`：PyQt 介面（分頁：產生、掃描、設定）
- `app/qr_tools.py`：QR 產生與資料處理
- `app/google_sheets.py`：Google Sheet 讀寫
- `app/config.py`：設定檔讀寫（預設 `./setting/config.json`）
- `app/paths.py`：集中管理 `./setting` 與 `./client` 路徑（支援打包）

---
若你需要將其包成單一 EXE（免安裝 Python），可用 PyInstaller：

```bash
.\.venv\Scripts\pip install pyinstaller
.\.venv\Scripts\pyinstaller -F -w -n QRcode_Signin app/main.py
```

打包完成的執行檔位於 `dist/QRcode_Signin.exe`。
