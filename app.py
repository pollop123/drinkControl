import os
import json
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from google.oauth2 import service_account
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials

app = Flask(__name__)

# LINE Bot的設定
LINE_CHANNEL_ACCESS_TOKEN = os.getenv('LINE_CHANNEL_ACCESS_TOKEN')
LINE_CHANNEL_SECRET = os.getenv('LINE_CHANNEL_SECRET')

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# Google Sheets API的設定
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
GOOGLE_SHEETS_API_KEY = os.getenv('GOOGLE_SHEETS_API_KEY')

if GOOGLE_SHEETS_API_KEY:
    credentials = Credentials.from_authorized_user_info(
        {"token": None, "scopes": SCOPES, "expiry": None},
        api_key=GOOGLE_SHEETS_API_KEY
    )
else:
    raise ValueError("No Google Sheets API key found in environment variables")

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
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="請提供您的 Google Sheet 連結")
        )
        # 等待用户提供链接，将链接解析并存储
        user_sheets[user_id] = None
    elif user_sheets[user_id] is None:
        # 保存用户提供的 Google Sheet 链接
        user_sheets[user_id] = user_message.split('/')[-2]  # 假设用户直接提供完整的 Google Sheet 链接
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="已成功連結到您的 Google Sheet。請輸入 '新增' 開始新增資料")
        )
    elif user_message == '新增':
        user_input_stage[user_id] = 'category'  # 初始化输入阶段
        user_input_data[user_id] = {}  # 初始化输入数据
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="請輸入種類")
        )
    else:
        # 根据当前的输入阶段来处理用户的输入
        stage = user_input_stage.get(user_id, None)
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
            # 所有数据已收集完毕，将其保存到 Google Sheets
            try:
                append_values(user_sheets[user_id], user_input_data[user_id])
                response_message = "訊息已儲存到Google Sheets"
            except Exception as e:
                response_message = f"無法儲存訊息到Google Sheets: {str(e)}"
            
            # 重置用户输入阶段
            user_input_stage[user_id] = None
            user_input_data[user_id] = {}

            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=response_message)
            )
        else:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="無法識別的命令或輸入。請輸入 '新增' 開始新增資料")
            )

def append_values(spreadsheet_id, data):
    # 准备数据
    values = [[data['category'], data['name'], data['calories']]]
    body = {
        'values': values
    }

    # 获取 A 列的所有数据
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range='Sheet1!A1:C'
    ).execute()
    rows = result.get('values', [])

    # 计算下一个空白单元格的位置
    next_row = len(rows) + 1
    next_cell = f'Sheet1!A{next_row}'

    # 调用 Google Sheets API 追加数据
    result = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=next_cell,
        valueInputOption='RAW',
        body=body
    ).execute()

    print(f"成功追加資料到 {next_cell}")

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
