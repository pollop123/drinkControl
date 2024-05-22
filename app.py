from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
import os.path
from google.oauth2 import service_account
import google.auth
from googleapiclient.discovery import build
from flask import Flask, request, abort
import re

# 設定Line Bot的Channel Secret和Access Token
line_bot_api = LineBotApi('Or51MZMEpmO44ahCt4PjwygnkbY76Pepve3pmoCUrj2qwyBcfKz+OlLCpsR8WZ8nIY5NPCdY3aXEKq8uQ2OJObUs6x52RwlXRwwJx3Jma4NbE5q/OzwVI1S9UYxavAEIz9dFs0DkW/w4DMHk7/uw/wdB04t89/1O/w1cDnyilFU=')
handler = WebhookHandler('626f45744c18177d7e4c0b0934e8f16c')

# 設定Google Sheets API的認證
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'drinking-control-fd0b44375d5c.json'

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=credentials)

# 設定Flask應用
app = Flask(__name__)

# 儲存用戶的Google Sheets ID
user_sheets = {}

@app.route("/callback", methods=['POST'])
def callback():
    # 獲取Line的請求體
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_id = event.source.user_id
    user_message = event.message.text

    # 檢查是否是Google Sheets連結
    sheet_id = extract_sheet_id(user_message)
    if sheet_id:
        user_sheets[user_id] = sheet_id
        reply_text = "已記錄您的Google Sheets連結！"
    else:
        if user_id in user_sheets:
            # 如果用戶已經提供了Sheets連結，將消息記錄到Sheet中
            SPREADSHEET_ID = user_sheets[user_id]
            RANGE_NAME = 'Sheet1!A1'  # 修改為你的Sheet和範圍
            value_input_option = 'RAW'
            values = [
                [user_message],
            ]
            body = {
                'values': values
            }
            result = service.spreadsheets().values().append(
                spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME,
                valueInputOption=value_input_option, body=body).execute()
            reply_text = "已將您的訊息保存到Google Sheets!"
        else:
            reply_text = "請先提供您的Google Sheets連結。"

    # 回應用戶
    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text=reply_text))

def extract_sheet_id(url):
    # 使用正則表達式提取Google Sheets ID
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if match:
        return match.group(1)
    return None

if __name__ == "__main__":
    app.run()
