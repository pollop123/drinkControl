# drinkControl

drinkControl 是以 Flask 建立的 LINE 飲食紀錄 Bot。使用者可以傳送食物或飲料照片，由 OpenAI Vision 辨識品項並估算熱量，也可以透過文字指令將飲食資料寫入個人的 Google Sheet。

## 功能

- 接收 LINE 圖片訊息並估算食物、飲料熱量
- 圖片送出前自動縮小，降低傳輸量與 API 成本
- 使用 OpenAI Responses API 回傳結構化 JSON
- 連結使用者提供的 Google Sheet
- 新增與刪除飲食紀錄
- 統計最近 1 天或 7 天的總熱量
- 計算各食物種類的熱量比例
- 查詢內建食物資料庫的熱量與份量

圖片辨識結果目前只會回覆到 LINE，不會自動寫入 Google Sheet。需要記錄資料時，請使用 `新增` 指令。

## 使用方式

第一次傳送文字訊息時，Bot 會要求提供 Google Sheet 連結。試算表必須授權給設定於 `GOOGLE_SERVICE_ACCOUNT_INFO` 的服務帳戶。

連結成功後可使用以下指令：

| 指令 | 說明 |
| --- | --- |
| `新增` | 依序輸入種類、名稱與熱量，新增一筆紀錄 |
| `清除` | 清除試算表內所有飲食紀錄 |
| `刪除上一筆` | 刪除最新一筆紀錄 |
| `加總` | 選擇最近 1 天或 7 天，計算總熱量 |
| `飲食比例` | 計算各種類占總熱量的比例 |
| `查詢大卡` | 查詢內建食物資料庫 |

也可以直接傳送圖片。Bot 會辨識照片中的主要食物與飲料，並回覆各品項及預估熱量。

## 技術架構

- Flask：處理 LINE webhook
- LINE Messaging API：接收訊息與傳送回覆
- OpenAI Responses API：圖片辨識與熱量估算
- Pillow：圖片方向修正、縮圖與 JPEG 壓縮
- Google Sheets API、gspread：儲存與統計飲食紀錄

圖片預設最長邊為 768 px，JPEG 品質為 88。圖片會保留原始比例，不裁切，也不會放大尺寸較小的圖片。

## 環境變數

複製 `.env.example` 為 `.env`，再填入實際設定：

```bash
cp .env.example .env
```

| 變數 | 必要 | 說明 |
| --- | --- | --- |
| `LINE_CHANNEL_ACCESS_TOKEN` | 是 | LINE Messaging API Channel Access Token |
| `LINE_CHANNEL_SECRET` | 是 | LINE Messaging API Channel Secret |
| `GOOGLE_SERVICE_ACCOUNT_INFO` | 是 | Google 服務帳戶完整 JSON，需放在同一行 |
| `OPENAI_API_KEY` | 是 | OpenAI API key |
| `CV_PROVIDER` | 是 | 目前僅支援 `openai` |
| `OPENAI_VISION_MODEL` | 否 | 圖片辨識模型，預設 `gpt-4.1-mini` |
| `OPENAI_IMAGE_DETAIL` | 否 | 圖片細節等級，預設 `low` |
| `IMAGE_MAX_SIZE` | 否 | 圖片最長邊，預設 `768` |
| `IMAGE_JPEG_QUALITY` | 否 | JPEG 品質，預設 `88` |

不要提交 `.env`。專案的 `.gitignore` 已排除 `.env`。

## 本機執行

建議使用虛擬環境：

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python app.py
```

本機服務預設位於 `http://127.0.0.1:5000`。若要接收 LINE webhook，需要使用可公開存取的 HTTPS 網址，並將 callback 設為：

```text
https://your-domain.example/callback
```

## Render 部署

在 Render 建立 Web Service，連結此 GitHub repository，並設定：

```text
Build Command: pip install -r requirements.txt
Start Command: gunicorn app:app
```

將 `.env.example` 中的變數逐項加入 Render Environment。部署完成後，把以下網址填入 LINE Developers Console 的 Webhook URL，並啟用 webhook：

```text
https://your-render-service.onrender.com/callback
```

使用免費方案時，服務休眠後的第一次請求可能因冷啟動而較慢。LINE webhook 對回應時間有限制，正式使用時應考慮避免服務休眠。

## 測試

先確認程式可以載入：

```bash
python -m py_compile app.py cv_analyzer.py image_utils.py
python -c "import app; print('app import ok')"
```

部署後可在 LINE Developers Console 執行 Verify，確認 webhook 回傳 `200 OK`。最後從 LINE 對 Bot 傳送一張包含食物或飲料的照片，確認完整流程：

```text
LINE 圖片訊息
-> Render webhook
-> 圖片壓縮
-> OpenAI 圖片辨識
-> LINE 回覆熱量結果
```

## 注意事項

- 圖片辨識與熱量為估算值，不適合作為醫療或營養診斷依據。
- 使用者的 Google Sheet 連結與輸入階段目前保存在應用程式記憶體中，服務重新啟動後需要重新連結。
- 建議在 OpenAI 平台設定專案預算與用量通知。
- 正式環境應檢查 Render logs，並監控 LINE webhook timeout 與 OpenAI API 錯誤。
