# DrinkControl 系統升級與優化指南 (Roadmap)

為解決目前遇到的「電腦視覺 (CV) 準確度低」與「Render 伺服器休眠導致綁定遺失」的兩大痛點，我們規劃了以下兩個階段的升級目標。本計畫遵守成本控制原則（不升級昂貴模型、不使用高耗能本地 YOLO 模型）。

---

## 階段一：CV 準確度與可信賴度爆發性升級
**核心策略**：AI 專注於特徵辨識，後端 (`food.py`) 專注於精準熱量計算。並透過「提示機制」引導 AI 發揮極限。

### 1. 修改 `cv_analyzer.py` (AI 分析大腦)
*   [ ] **提升影像解析度**：將 `os.getenv("OPENAI_IMAGE_DETAIL", "low")` 修改為 `"high"` 或是預設改為 `"auto"`，讓 AI 能看清楚食物細節。
*   [ ] **修改 JSON Schema (`FOOD_ANALYSIS_SCHEMA`)**：
    *   新增 `reasoning` 欄位：強制 AI 輸出結果前先「思考並記錄觀察特徵」(Chain-of-Thought)，大幅降低幻覺。
    *   移除讓 AI 猜測的 `calories`。
    *   新增 `estimated_weight_g` (預估重量) 與 `fallback_calories` (找不到時的備用熱量)。
*   [ ] **強化 Prompt (`ANALYSIS_PROMPT`)**：
    *   **導入 Few-Shot Prompting**：在 Prompt 中寫死 1~2 個完美範例，引導 AI 穩定輸出格式。
    *   **支援 Context Hint**：允許 `analyze_image` 函數接收額外的文字提示字串，動態插入 Prompt 中。

### 2. 修改 `app.py` (後端邏輯與 LINE Bot 互動)
*   [x] **完成圖片記錄主流程**：圖片辨識後顯示確認／取消按鈕，確認後寫入 Google Sheet 並回覆今日進度。
*   [x] **加入 LINE 快速按鈕**：常用功能改由 Quick Reply 操作，並保留文字指令相容性。
*   [x] **改善完整使用流程**：加入辨識熱量修改、重新拍照、分析中提示、個人目標與分期間飲食比例。
*   [x] **安全撤銷與清除**：以 `meal_id` 整餐撤銷，危險操作使用一次性限時確認。
*   [x] **常駐 Rich Menu**：提供可重複執行的建立工具與預覽圖產生器。
*   [x] **餐次卡片與歷史**：以 Flex 卡片查看今日、指定日期及最近餐次，並能編輯指定品項。
*   [x] **最近餐點再記一次**：保留品項與熱量，使用新的餐次及品項 ID 安全複製。
*   [x] **每週飲食報告**：提供隨選查看、前週比較、分類分析與使用者可選擇的自動推送。
*   [ ] **實作圖片與文字輔助 (Context Hint)**：
    *   新增全域變數 `user_hints = {}` 暫存使用者的文字提示。
    *   當使用者傳送非指令文字時（例如：「這是一個排骨便當」），暫存該文字並回覆：「已記錄提示，請上傳圖片進行分析」。
    *   收到圖片時，一併將 Hint 傳給 `analyze_image`。
*   [ ] **串接本地精準資料庫 (`food.py`)**：
    *   新增 `calculate_calories` 函式。收到 AI 辨識的品項名稱後，去 `food_dict` 進行模糊搜尋。
    *   有找到：用 `(查表熱量 * estimated_weight_g) / 100` 計算絕對精準的熱量。
    *   沒找到：使用 AI 的 `fallback_calories`。
*   [ ] **優化 LINE 回覆格式**：
    *   清楚標示資料來源，例如：`偵測到: 鮮奶茶 (約300g) -> 130 大卡 (資料庫)`，提升使用者信任感。

---

## 階段二：系統穩定性升級 (永久儲存使用者綁定)
**核心策略**：因為 Render 免費版的短暫檔案系統 (Ephemeral Filesystem) 每天會重啟並清空 RAM 與本地檔案，我們必須將 `user_sheets` 狀態轉移到外部。既然已經串好 Google Sheets API，最好的選擇就是用另一個獨立的 Google Sheet 當作 Config 資料庫。

### 1. 環境前置準備 (手動設定)
*   [ ] 在 Google 雲端硬碟建立一個新的試算表（例如命名為 `DrinkControl_Config`）。
*   [ ] 將該試算表的編輯權限分享給現有的 Google Service Account Email (寫在 `GOOGLE_SERVICE_ACCOUNT_INFO` 裡的那組)。
*   [ ] 複製該試算表的 ID，存到 Render 環境變數中，命名為 `CONFIG_SPREADSHEET_ID`。

### 2. 修改 `app.py` (資料庫存取)
*   [x] **初始化讀取**：程式啟動時，使用 `gspread` 開啟 `CONFIG_SPREADSHEET_ID`，並將試算表中的資料讀取出來，還原到 `user_sheets` 字典中。
*   [x] **動態寫入**：當遇到新的 `user_id` 綁定成功時，除了更新記憶體中的 `user_sheets`，還要將 `[user_id, spreadsheet_id]` 作為新的一列寫入 Config 試算表中。
*   [x] **保障機制**：這樣即便伺服器休眠重啟，也能在喚醒時自動找回所有使用者的綁定網址，徹底解決「需重複綁定」的痛點。
