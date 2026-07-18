# drinkControl

drinkControl 是以 Flask 建立的 LINE 飲食紀錄 Bot。使用者可以傳送食物或飲料照片，由 OpenAI Vision 辨識品項並估算熱量，確認後寫入個人的 Google Sheet。

## 功能

- 接收 LINE 圖片訊息並估算食物、飲料熱量
- 圖片送出前自動縮小，降低傳輸量與 API 成本
- 使用 OpenAI Responses API 回傳結構化 JSON
- 連結使用者提供的 Google Sheet
- 照片辨識後以按鈕確認或取消，自動新增飲食紀錄
- 辨識結果可在儲存前修改熱量或重新拍照
- 透過 LINE 快速按鈕新增、刪除與查詢紀錄
- 可安裝常駐 LINE Rich Menu
- 顯示今日熱量進度與最近 7 日總熱量
- 每位使用者可設定自己的每日熱量目標
- 以餐次識別碼整餐撤銷，不會只刪除照片中的最後一個品項
- 以 LINE Flex 卡片查看今日、指定日期與最近餐次
- 從餐次卡片修改指定品項的名稱、分類或熱量
- 將最近餐點「再記一次」，不需重新辨識照片
- 產生近 7 日週報，並可選擇開啟自動推送
- 統計最近 1 天或 7 天的總熱量
- 計算各食物種類的熱量比例
- 查詢內建食物資料庫的熱量與份量

圖片辨識完成後，Bot 會列出品項、預估熱量與可信度。按下「確認記錄」後，各品項會寫入 Google Sheet，並立即顯示今日累積熱量。

## 使用方式

第一次使用時，Bot 會引導使用者建立 Google Sheet、分享給 `GOOGLE_SERVICE_ACCOUNT_INFO` 中的服務帳戶 Email，再貼上試算表連結。連結完成後可設定個人每日熱量目標，也可略過。

連結成功後，Bot 會在訊息下方顯示功能按鈕。原本文字指令仍可使用：

| 指令 | 說明 |
| --- | --- |
| `新增` | 依序輸入種類、名稱與熱量，新增一筆紀錄 |
| `拍照說明` | 顯示拍照提示與開啟相機按鈕 |
| `今日` | 查看今日累積熱量與目標進度 |
| `近7日` | 查看最近 7 日總熱量 |
| `今日餐次` | 以卡片查看今天每一餐並編輯、再記或刪除 |
| `歷史紀錄` | 查看昨天、近 7 日或指定日期的餐次 |
| `最近餐點` | 查看最近五餐並可直接再記一次 |
| `每週報告` | 查看熱量趨勢、記錄天數與分類比例 |
| `設定` | 設定個人的每日熱量目標 |
| `清除` | 限時二次確認後清除試算表內所有飲食紀錄 |
| `刪除上一筆` | 預覽並撤銷上一餐的全部品項 |
| `加總` | 選擇最近 1 天或 7 天，計算總熱量 |
| `飲食比例` | 選擇今日、近 7 日或全部期間後計算種類比例 |
| `查詢大卡` | 查詢內建食物資料庫 |

也可以直接傳送圖片。Bot 會先回覆「正在分析」，接著列出主要食物、飲料、可信度與熱量。使用者可以確認、修改熱量、重新拍照或取消。

飲食 Sheet 使用以下欄位：`timestamp`、`category`、`name`、`calories`、`meal_id`、`item_id`。`meal_id` 用來識別整餐，`item_id` 用來安全編輯指定品項。舊版四欄或五欄 Sheet 會在下次操作時原地補齊識別欄，不會搬動既有資料。

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
| `CONFIG_SPREADSHEET_ID` | 是 | 保存 LINE 使用者與飲食 Sheet 綁定的 Config Sheet ID |
| `OPENAI_API_KEY` | 是 | OpenAI API key |
| `CV_PROVIDER` | 是 | 目前僅支援 `openai` |
| `OPENAI_VISION_MODEL` | 否 | 圖片辨識模型，預設 `gpt-4.1-mini` |
| `OPENAI_IMAGE_DETAIL` | 否 | 圖片細節等級，預設 `low` |
| `IMAGE_MAX_SIZE` | 否 | 圖片最長邊，預設 `768` |
| `IMAGE_JPEG_QUALITY` | 否 | JPEG 品質，預設 `88` |
| `DAILY_CALORIE_TARGET` | 否 | 尚未設定個人目標時的預設值；`0` 代表只顯示累積 |
| `PENDING_ACTION_TTL_MINUTES` | 否 | 清除、撤銷與圖片確認按鈕的有效分鐘數，預設 `10` |
| `CRON_SECRET` | 自動週報需要 | 保護 `/cron/weekly-report` 的長隨機密鑰 |

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

## LINE Rich Menu

先產生預覽確認排版：

```bash
python setup_rich_menu.py --preview rich-menu-preview.png
```

確認後，以 `.env` 中的 LINE token 建立並設為預設 Rich Menu：

```bash
python setup_rich_menu.py
```

工具會重用同名 Rich Menu，不會因重複執行持續建立副本。新版選單包含拍照記錄、今日餐次、歷史紀錄、最近餐點、每週報告與個人設定。

## 自動每週報告

使用者可在「個人設定」開啟或關閉自動週報。設定 `CRON_SECRET` 後，每週由 Render Cron Job 呼叫：

```bash
curl -X POST https://your-render-service.onrender.com/cron/weekly-report \
  -H "Authorization: Bearer $CRON_SECRET"
```

建議排程在每週日晚上。端點只會推送給已主動開啟週報且近 7 日有紀錄的使用者；未帶正確密鑰會回傳 `401`。

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
python -m py_compile app.py cv_analyzer.py image_utils.py meal_views.py setup_rich_menu.py
python -c "import app; print('app import ok')"
python -m unittest discover -s tests -v
```

部署後可在 LINE Developers Console 執行 Verify，確認 webhook 回傳 `200 OK`。最後從 LINE 對 Bot 傳送一張包含食物或飲料的照片，確認完整流程：

```text
LINE 圖片訊息
-> Render webhook
-> 圖片壓縮
-> OpenAI 圖片辨識
-> LINE 顯示確認按鈕
-> 確認後寫入 Google Sheet
-> LINE 回覆今日進度
```

## 注意事項

- 圖片辨識與熱量為估算值，不適合作為醫療或營養診斷依據。
- 使用者的 Google Sheet 綁定與個人熱量目標保存在 Config Sheet。
- 尚未確認的圖片與輸入階段只存在應用程式記憶體中；服務重啟後舊按鈕會安全失效，不會寫入或刪除資料。
- 建議在 OpenAI 平台設定專案預算與用量通知。
- 正式環境應檢查 Render logs，並監控 LINE webhook timeout 與 OpenAI API 錯誤。
