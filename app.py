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

# 全局變量來存儲 Google Sheet ID
user_sheets = {}

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
        # 將 Google Sheet ID 提取出來並存儲
        user_sheets[user_id] = user_message.split('/')[-2]  # 假設用戶直接提供完整的 Google Sheet 链接
    else:
        # 用戶已經提供了 Google Sheet 連結，將訊息追加到相應的表格中
        try:
            append_values(user_sheets[user_id], user_message)
            response_message = "訊息已儲存到Google Sheets"
        except Exception as e:
            response_message = f"無法儲存訊息到Google Sheets: {str(e)}"

        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=response_message)
        )

def append_values(spreadsheet_id, value):
    # 準備資料
    values = [[value]]
    body = {
        'values': values
    }

    # 獲取 A 列的下一個空白儲存格位置
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range='Sheet1!A:A'
    ).execute()
    next_row = len(result.get('values', [])) + 1

    # 呼叫 Google Sheets API 追加資料
    result = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=f'Sheet1!A{next_row}',
        valueInputOption='RAW',
        body=body
    ).execute()


if __name__ == "__main__":
    app.run()
