import os
import json
from dotenv import load_dotenv

load_dotenv()
import re
import openpyxl
from datetime import datetime, timedelta
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage, ImageMessage
from google.oauth2.service_account import Credentials
import gspread
from cv_analyzer import analyze_image
from image_utils import compress_image
from food import food_dict

app = Flask(__name__)

# LINE Bot設定
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
IMAGE_MAX_SIZE = int(os.getenv("IMAGE_MAX_SIZE", "768"))
IMAGE_JPEG_QUALITY = int(os.getenv("IMAGE_JPEG_QUALITY", "88"))

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# Google Sheets API設定
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SERVICE_ACCOUNT_INFO = os.getenv("GOOGLE_SERVICE_ACCOUNT_INFO")
CONFIG_SPREADSHEET_ID = os.getenv("CONFIG_SPREADSHEET_ID")

if SERVICE_ACCOUNT_INFO:
    service_account_info = json.loads(SERVICE_ACCOUNT_INFO)
    credentials = Credentials.from_service_account_info(
        service_account_info, scopes=SCOPES
    )
else:
    raise ValueError("環境變數中找不到 Google 服務帳戶資訊")

gc = gspread.authorize(credentials)

# 全域變數存儲用戶輸入狀態
user_sheets = {}
user_input_stage = {}
user_input_data = {}
config_sheet = None

CONFIG_HEADERS = ["user_id", "spreadsheet_id", "updated_at"]


def get_config_sheet():
    global config_sheet

    if not CONFIG_SPREADSHEET_ID:
        raise ValueError("環境變數中找不到 CONFIG_SPREADSHEET_ID")

    if config_sheet is None:
        config_sheet = gc.open_by_key(CONFIG_SPREADSHEET_ID).sheet1

    headers = config_sheet.row_values(1)
    if not headers:
        config_sheet.append_row(CONFIG_HEADERS)
    elif headers != CONFIG_HEADERS:
        raise ValueError(
            "Config Sheet 標題列必須是 user_id、spreadsheet_id、updated_at"
        )

    return config_sheet


def load_user_sheet_bindings():
    if not CONFIG_SPREADSHEET_ID:
        app.logger.warning("CONFIG_SPREADSHEET_ID 未設定，使用者綁定只會保存在記憶體")
        return

    try:
        sheet = get_config_sheet()
        rows = sheet.get_all_values()
        for row in rows[1:]:
            if len(row) >= 2 and row[0] and row[1]:
                user_sheets[row[0]] = row[1]
        app.logger.info("已載入 %s 筆使用者 Sheet 綁定", len(user_sheets))
    except Exception as e:
        app.logger.error("無法載入使用者 Sheet 綁定: %s", e)


def save_user_sheet_binding(user_id, spreadsheet_id):
    sheet = get_config_sheet()
    rows = sheet.get_all_values()
    updated_at = (datetime.utcnow() + timedelta(hours=8)).strftime(
        "%Y-%m-%d %H:%M:%S"
    )
    values = [user_id, spreadsheet_id, updated_at]

    for row_number, row in enumerate(rows[1:], start=2):
        if row and row[0] == user_id:
            sheet.update(
                values=[values],
                range_name=f"A{row_number}:C{row_number}",
            )
            return

    sheet.append_row(values)


load_user_sheet_bindings()


# 解析 Google Sheet 連結以獲取 spreadsheetId
def get_spreadsheet_id(url):
    pattern = r"\/spreadsheets\/d\/([a-zA-Z0-9-_]+)"
    match = re.search(pattern, url)
    if match:
        return match.group(1)
    else:
        return None


@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers["X-Line-Signature"]
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return "OK"


@handler.add(MessageEvent, message=ImageMessage)
def handle_image_message(event):
    message_id = event.message.id
    temp_image_path = f"{message_id}.jpg"

    try:
        message_content = line_bot_api.get_message_content(message_id)
        with open(temp_image_path, "wb") as fd:
            for chunk in message_content.iter_content():
                fd.write(chunk)

        compress_image(
            temp_image_path,
            max_size=IMAGE_MAX_SIZE,
            quality=IMAGE_JPEG_QUALITY,
        )
        analysis_result = analyze_image(temp_image_path)

        if analysis_result and analysis_result.get("detections"):
            detections = analysis_result["detections"]
            if len(detections) == 1:
                d = detections[0]
                response_message = f"偵測到: {d['item']}\n熱量約: {d['calories']} 大卡"
            else:
                lines = ["偵測到多種食物:"]
                total = 0
                for d in detections:
                    lines.append(f"• {d['item']}: {d['calories']} 大卡")
                    total += d["calories"]
                lines.append(f"\n總熱量約: {total} 大卡")
                response_message = "\n".join(lines)
        else:
            response_message = "抱歉，無法辨識圖片中的食物或飲料。"

        line_bot_api.reply_message(
            event.reply_token, TextSendMessage(text=response_message)
        )

    except Exception as e:
        app.logger.error(f"Error processing image: {e}")
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="處理圖片時發生錯誤，請稍後再試。"),
        )
    finally:
        if os.path.exists(temp_image_path):
            os.remove(temp_image_path)


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_id = event.source.user_id
    user_message = event.message.text.strip()

    if user_id not in user_sheets:
        user_sheets[user_id] = None

    if user_sheets[user_id] is None:
        spreadsheet_id = get_spreadsheet_id(user_message)
        if not spreadsheet_id:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="請提供有效的 Google Sheet 連結"),
            )
            return

        try:
            add_headers(spreadsheet_id)
            save_user_sheet_binding(user_id, spreadsheet_id)
        except Exception as e:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=f"連結失敗或無法保存綁定: {str(e)}"),
            )
            return

        user_sheets[user_id] = spreadsheet_id
        line_bot_api.reply_message(
            event.reply_token,
            [
                TextSendMessage(text="已成功連結到您的 Google Sheet。"),
                TextSendMessage(
                    text="您現在可以傳送圖片來分析熱量，或輸入以下指令：\n'新增'：開始新增資料\n'清除'：刪除所有資料\n'刪除上一筆'：刪除上一筆新增資料\n'加總'：加總大卡\n'飲食比例'：獲取各種類的熱量比例\n'查詢大卡'：獲取各種食物的熱量"
                ),
            ],
        )
        return

    if user_message == "新增":
        user_input_stage[user_id] = "category"
        user_input_data[user_id] = {}
        line_bot_api.reply_message(
            event.reply_token, TextSendMessage(text="請輸入種類")
        )
        return

    if user_message == "清除":
        try:
            clear_sheet(user_sheets[user_id])
            response_message = "已清除所有資料"
        except Exception as e:
            response_message = f"無法清除資料: {str(e)}"
        line_bot_api.reply_message(
            event.reply_token, TextSendMessage(text=response_message)
        )
        return

    if user_message == "刪除上一筆":
        try:
            delete_last_entry(user_sheets[user_id])
            response_message = "已刪除最新的一筆資料"
        except Exception as e:
            response_message = f"無法刪除資料: {str(e)}"
        line_bot_api.reply_message(
            event.reply_token, TextSendMessage(text=response_message)
        )
        return

    if user_message == "加總":
        user_input_stage[user_id] = "sum_period"
        line_bot_api.reply_message(
            event.reply_token, TextSendMessage(text="請輸入 '1天' 或 '7天' 來加總大卡")
        )
        return

    if user_message == "飲食比例":
        try:
            category_ratios = calculate_category_ratios(user_sheets[user_id])
            response_message = "飲食比例:\n" + "\n".join(
                [
                    f"{category}: {ratio:.2f}%"
                    for category, ratio in category_ratios.items()
                ]
            )
        except Exception as e:
            response_message = f"無法計算飲食比例: {str(e)}"
        line_bot_api.reply_message(
            event.reply_token, TextSendMessage(text=response_message)
        )
        return

    if user_message == "查詢大卡":
        user_input_stage[user_id] = "query_kcal"
        line_bot_api.reply_message(
            event.reply_token, TextSendMessage(text="請輸入要查詢的食物名稱")
        )
        return

    stage = user_input_stage.get(user_id)

    if stage == "query_kcal":
        result = search_kcal(user_message)
        response_message = result or f"抱歉，我們的資料庫中沒有 '{user_message}' 的熱量資訊"
        user_input_stage[user_id] = None
        line_bot_api.reply_message(
            event.reply_token, TextSendMessage(text=response_message)
        )
        return

    if stage == "category":
        user_input_data.setdefault(user_id, {})["category"] = user_message
        user_input_stage[user_id] = "name"
        line_bot_api.reply_message(
            event.reply_token, TextSendMessage(text="請輸入名稱")
        )
        return

    if stage == "name":
        user_input_data.setdefault(user_id, {})["name"] = user_message
        user_input_stage[user_id] = "calories"
        line_bot_api.reply_message(
            event.reply_token, TextSendMessage(text="請輸入大卡")
        )
        return

    if stage == "calories":
        try:
            calories = int(user_message)
            if calories < 0:
                raise ValueError
        except ValueError:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="大卡必須是大於或等於 0 的整數，請重新輸入"),
            )
            return

        user_input_data.setdefault(user_id, {})["calories"] = calories
        user_input_data[user_id]["timestamp"] = (
            datetime.utcnow() + timedelta(hours=8)
        ).strftime("%Y-%m-%d %H:%M:%S")

        try:
            append_values(user_sheets[user_id], user_input_data[user_id])
            response_message = "訊息已儲存到 Google Sheets"
        except Exception as e:
            response_message = f"無法儲存訊息到 Google Sheets: {str(e)}"

        user_input_stage[user_id] = None
        user_input_data[user_id] = {}
        line_bot_api.reply_message(
            event.reply_token, TextSendMessage(text=response_message)
        )
        return

    if stage == "sum_period":
        if user_message not in ["1天", "7天"]:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="無效的輸入，請輸入 '1天' 或 '7天'"),
            )
            return

        try:
            days = 1 if user_message == "1天" else 7
            total_calories = sum_calories(user_sheets[user_id], days)
            response_message = f"{user_message} 的總大卡為 {total_calories} 大卡"
        except Exception as e:
            response_message = f"無法計算總大卡: {str(e)}"

        user_input_stage[user_id] = None
        line_bot_api.reply_message(
            event.reply_token, TextSendMessage(text=response_message)
        )
        return

    line_bot_api.reply_message(
        event.reply_token,
        [
            TextSendMessage(text="無法識別的命令"),
            TextSendMessage(
                text="請輸入 '新增' 開始新增資料，輸入 '清除' 刪除所有資料，輸入 '刪除上一筆' 刪除上一筆新增資料，輸入 '加總' 加總大卡，輸入 '飲食比例' 獲取各種類的熱量比例，輸入 '查詢大卡' 獲取各種食物的熱量"
            ),
        ],
    )


def search_kcal(name):
    for key in food_dict.keys():
        if name in key:
            return f"品項：{key}, 熱量：{int(food_dict[key][0])} kcal, 份量：{food_dict[key][1]}"
    return None


def add_headers(spreadsheet_id):
    sheet = gc.open_by_key(spreadsheet_id).sheet1
    headers = ["timestamp", "category", "name", "calories"]
    if sheet.row_values(1) != headers:
        sheet.insert_row(headers, 1)
        print(f"已添加標題行到試算表 {spreadsheet_id}")


def append_values(spreadsheet_id, data):
    sheet = gc.open_by_key(spreadsheet_id).sheet1
    row = [data["timestamp"], data["category"], data["name"], data["calories"]]
    sheet.append_row(row)
    print(f"成功將資料 {row} 新增到試算表 {spreadsheet_id}")


def clear_sheet(spreadsheet_id):
    sheet = gc.open_by_key(spreadsheet_id).sheet1
    sheet.clear()
    sheet.append_row(["timestamp", "category", "name", "calories"])
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
    total_calories = sum(
        int(record["calories"])
        for record in all_records
        if (now - datetime.strptime(record["timestamp"], "%Y-%m-%d %H:%M:%S")).days
        < days
    )
    return total_calories


def calculate_category_ratios(spreadsheet_id):
    sheet = gc.open_by_key(spreadsheet_id).sheet1
    all_records = sheet.get_all_records()
    category_totals = {}
    total_calories = 0

    for record in all_records:
        category = record["category"]
        calories = int(record["calories"])
        total_calories += calories
        if category in category_totals:
            category_totals[category] += calories
        else:
            category_totals[category] = calories

    category_ratios = {
        category: (calories / total_calories) * 100
        for category, calories in category_totals.items()
    }
    return category_ratios


if __name__ == "__main__":
    app.run(debug=True)
