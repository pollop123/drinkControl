import os
import json
import re
import openpyxl
from datetime import datetime, timedelta
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
    raise ValueError("環境變數中找不到 Google 服務帳戶資訊")

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
            try:
                add_headers(spreadsheet_id)
                line_bot_api.reply_message(
                    event.reply_token,
                    [
                        TextSendMessage(text="已成功連結到您的 Google Sheet。"),
                        TextSendMessage(text="請輸入 '新增' 開始新增資料，輸入 '清除' 刪除所有資料，輸入 '刪除上一筆' 刪除上一筆新增資料，輸入 '加總' 加總大卡，輸入 '飲食比例' 獲取各種類食物的熱量比例")
                    ]
                )
            except Exception as e:
                line_bot_api.reply_message(
                    event.reply_token,
                    TextSendMessage(text=f"連結失敗或無法添加標題行: {str(e)}")
                )
        else:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="無效的 Google Sheet 連結，請重新提供")
            )

    elif user_message == '新增':
        user_input_stage[user_id] = 'category'
        user_input_data[user_id] = {}
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="請輸入種類")
        )

    elif user_message == '清除':
        try:
            clear_sheet(user_sheets[user_id])
            response_message = "已清除所有資料"
        except Exception as e:
            response_message = f"無法清除資料: {str(e)}"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=response_message)
        )

    elif user_message == '刪除上一筆':
        try:
            delete_last_entry(user_sheets[user_id])
            response_message = "已刪除最新的一筆資料"
        except Exception as e:
            response_message = f"無法刪除資料: {str(e)}"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=response_message)
        )

    elif user_message == '加總':
        user_input_stage[user_id] = 'sum_period'
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="請輸入 '1天' 或 '7天' 來加總大卡")
        )

    elif user_message == '飲食比例':
        try:
            category_ratios = calculate_category_ratios(user_sheets[user_id])
            response_message = "飲食比例:\n" + "\n".join([f"{category}: {ratio:.2f}%" for category, ratio in category_ratios.items()])
        except Exception as e:
            response_message = f"無法計算飲食比例: {str(e)}"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=response_message)
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
            user_input_data[user_id]['timestamp'] = (datetime.utcnow() + timedelta(hours=8)).strftime('%Y-%m-%d %H:%M:%S')
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
        elif stage == 'sum_period':
            try:
                if user_message in ['1天', '7天']:
                    days = 1 if user_message == '1天' else 7
                    total_calories = sum_calories(user_sheets[user_id], days)
                    response_message = f"{user_message} 的總大卡為 {total_calories} 大卡"
                else:
                    response_message = "無效的輸入，請輸入 '1天' 或 '7天'"
            except Exception as e:
                response_message = f"無法計算總大卡: {str(e)}"
            
            user_input_stage[user_id] = None
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=response_message)
            )
        else:
            line_bot_api.reply_message(
                event.reply_token,
                [
                    TextSendMessage(text="無法識別的命令"),
                    TextSendMessage(text="請輸入 '新增' 開始新增資料，輸入 '清除' 刪除所有資料，輸入 '刪除上一筆' 刪除上一筆新增資料，輸入 '加總' 加總大卡，輸入 '飲食比例' 獲取各種類食物的熱量比例")
                ]
            )

def add_headers(spreadsheet_id):
    sheet = gc.open_by_key(spreadsheet_id).sheet1
    headers = ['timestamp', 'category', 'name', 'calories']
    if sheet.row_values(1) != headers:
        sheet.insert_row(headers, 1)
        print(f"已添加標題行到試算表 {spreadsheet_id}")

def append_values(spreadsheet_id, data):
    sheet = gc.open_by_key(spreadsheet_id).sheet1
    row = [data['timestamp'], data['category'], data['name'], data['calories']]
    sheet.append_row(row)
    print(f"成功將資料 {row} 新增到試算表 {spreadsheet_id}")

def clear_sheet(spreadsheet_id):
    sheet = gc.open_by_key(spreadsheet_id).sheet1
    sheet.clear()
    sheet.append_row(['timestamp', 'category', 'name', 'calories'])
    print(f"已清除試算表 {spreadsheet_id} 的所有資料")

def delete_last_entry(spreadsheet_id):
    sheet = gc.open_by_key(spreadsheet_id).sheet1
    last_row = len(sheet.get_all_values())
    if last_row > 1:
        sheet.delete_rows(last_row)
        print(f"已刪除試算表 {spreadsheet_id} 的最新一筆資料")
    else:
        print("試算表中無可刪除的資料")

def sum_calories(spreadsheet_id, days):
    sheet = gc.open_by_key(spreadsheet_id).sheet1
    all_records = sheet.get_all_records()
    now = datetime.utcnow() + timedelta(hours=8)
    total_calories = sum(int(record['calories']) for record in all_records if (now - datetime.strptime(record['timestamp'], '%Y-%m-%d %H:%M:%S')).days < days)
    return total_calories

def calculate_category_ratios(spreadsheet_id):
    sheet = gc.open_by_key(spreadsheet_id).sheet1
    all_records = sheet.get_all_records()
    category_totals = {}
    total_calories = 0
    
    for record in all_records:
        category = record['category']
        calories = int(record['calories'])
        total_calories += calories
        if category in category_totals:
            category_totals[category] += calories
        else:
            category_totals[category] = calories
    
    category_ratios = {category: (calories / total_calories) * 100 for category, calories in category_totals.items()}
    return category_ratios

if __name__ == "__main__":
    app.run(debug=True)
