#嗨
import os
import json
import re
import openpyxl
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from google.oauth2.service_account import Credentials
import gspread

app = Flask(__name__)

# LINE Bot設定
LINE_CHANNEL_ACCESS_TOKEN = os.getenv('LINE_CHANNEL_ACCESS_TOKEN')
LINE_CHANNEL_SECRET = os.getenv('LINE_CHANNEL_SECRET')

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# Google Sheets API設定
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_INFO = os.getenv('GOOGLE_SERVICE_ACCOUNT_INFO')

if SERVICE_ACCOUNT_INFO:
    service_account_info = json.loads(SERVICE_ACCOUNT_INFO)
    credentials = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
else:
    raise ValueError("No Google service account info found in environment variables")

gc = gspread.authorize(credentials)

# 全域變數存儲用戶輸入狀態
user_sheets = {}
user_input_stage = {}
user_input_data = {}

# 解析 Google Sheet 連結以獲取 spreadsheetId
def get_spreadsheet_id(url):
    pattern = r'\/spreadsheets\/d\/([a-zA-Z0-9-_]+)'
    match = re.search(pattern, url)
    if match:
        return match.group(1)
    else:
        return None

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

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
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="請提供您的 Google Sheet 連結")
        )
        user_sheets[user_id] = None

    elif user_sheets[user_id] is None:
        spreadsheet_id = get_spreadsheet_id(user_message)
        if spreadsheet_id:
            user_sheets[user_id] = spreadsheet_id
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="已成功連結到您的 Google Sheet。請輸入 '新增' 開始新增資料")
            )
        else:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="無效的 Google Sheet 連結,請重新提供")
            )

    elif user_message == '新增':
        user_input_stage[user_id] = 'category'
        user_input_data[user_id] = {}
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="請輸入種類")
        )

    else:
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
            try:
                append_values(user_sheets[user_id], user_input_data[user_id])
                response_message = "訊息已儲存到Google Sheets"
            except Exception as e:
                response_message = f"無法儲存訊息到Google Sheets: {str(e)}"
            
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
    # 使用 gspread 開啟指定的試算表
    sheet = gc.open_by_key(spreadsheet_id).sheet1

    # 將資料新增到試算表
    row = [data['category'], data['name'], data['calories']]
    sheet.append_row(row)
    print(f"成功將資料 {row} 新增到試算表 {spreadsheet_id}")

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
