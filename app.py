import os
import json
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from google.oauth2 import service_account
from googleapiclient.discovery import build

app = Flask(__name__)

# LINE Bot的設定
LINE_CHANNEL_ACCESS_TOKEN = os.getenv('LINE_CHANNEL_ACCESS_TOKEN')
LINE_CHANNEL_SECRET = os.getenv('LINE_CHANNEL_SECRET')

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# Google Sheets API的設定
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_INFO = os.getenv('GOOGLE_SERVICE_ACCOUNT_INFO')

if SERVICE_ACCOUNT_INFO:
    service_account_info = json.loads(SERVICE_ACCOUNT_INFO)
    credentials = service_account.Credentials.from_service_account_info(
        service_account_info, scopes=SCOPES)
else:
    raise ValueError("No Google service account info found in environment variables")

service = build('sheets', 'v4', credentials=credentials)

# 全局變量來存儲 Google Sheet ID 和用戶的輸入狀態
user_sheets = {}
user_input_stage = {}
user_input_data = {}

@app.route("/callback", methods=['POST'])
def callback():
    # 取得請求的簽名
    signature = request.headers['X-Line-Signature']

    # 取得請求的body
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    # 驗證簽名
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_id = event.source.user_id
    user_message = event.message.text

    if user_id not in user_sheets:
        # 如果用戶還沒有提供 Google Sheet 链接，請求他們提供
        user_sheets[user_id] = user_message.split('/')[-2]  # 假設用戶直接提供完整的 Google Sheet 链接
        user_input_stage[user_id] = 'category'  # 初始化输入阶段
        user_input_data[user_id] = {}  # 初始化输入数据
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="已成功連結到您的 Google Sheet。請輸入種類")
        )
    else:
        # 根據當前的輸入階段來處理用戶的輸入
        stage = user_input_stage.get(user_id, 'category')
        if stage == 'category':
            user_input_data[user_id]['category'] = user_message
            user_input_stage[user_id] = 'name'
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="請輸入名稱")
            )
        elif stage == 'name':
            user_input_data[user_id]['name'] = user_message
            user_input_stage[user_id] = 'calories'
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="請輸入大卡")
            )
        elif stage == 'calories':
            user_input_data[user_id]['calories'] = user_message
            # 所有數據已收集完畢，將其保存到 Google Sheets
            try:
                append_values(user_sheets[user_id], user_input_data[user_id])
                response_message = "訊息已儲存到Google Sheets"
            except Exception as e:
                response_message = f"無法儲存訊息到Google Sheets: {str(e)}"
            
            # 重置用户输入阶段
            user_input_stage[user_id] = 'category'
            user_input_data[user_id] = {}

            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=response_message)
            )

def append_values(spreadsheet_id, data):
    # 準備資料
    values = [[data['category'], data['name'], data['calories']]]
    body = {
        'values': values
    }

    # 獲取 A 列的所有資料
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range='Sheet1!A1:C'
    ).execute()
    rows = result.get('values', [])

    # 計算下一個空白儲存格的位置
    next_row = len(rows) + 1
    next_cell = f'Sheet1!A{next_row}'

    # 呼叫 Google Sheets API 追加資料
    result = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=next_cell,
        valueInputOption='RAW',
        body=body
    ).execute()

    print(f"成功追加資料到 {next_cell}")

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
