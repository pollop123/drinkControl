import os
import json
import secrets
import uuid
from dotenv import load_dotenv
import re
from datetime import datetime, timedelta, timezone
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import (
    ButtonsTemplate,
    CameraAction,
    ImageMessage,
    MessageAction,
    MessageEvent,
    PostbackAction,
    PostbackEvent,
    QuickReply,
    QuickReplyButton,
    TemplateSendMessage,
    TextMessage,
    TextSendMessage,
)
from google.oauth2.service_account import Credentials
import gspread
from cv_analyzer import analyze_image
from image_utils import compress_image
from food import food_dict

load_dotenv()

app = Flask(__name__)

# LINE Bot設定
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
IMAGE_MAX_SIZE = int(os.getenv("IMAGE_MAX_SIZE", "768"))
IMAGE_JPEG_QUALITY = int(os.getenv("IMAGE_JPEG_QUALITY", "88"))
DAILY_CALORIE_TARGET = int(os.getenv("DAILY_CALORIE_TARGET", "0"))
PENDING_ACTION_TTL_MINUTES = int(os.getenv("PENDING_ACTION_TTL_MINUTES", "10"))

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
user_targets = {}
user_input_stage = {}
user_input_data = {}
pending_image_records = {}
pending_clear_actions = {}
pending_delete_actions = {}
config_sheet = None

CONFIG_HEADERS = [
    "user_id",
    "spreadsheet_id",
    "updated_at",
    "daily_calorie_target",
]
LEGACY_CONFIG_HEADERS = ["user_id", "spreadsheet_id", "updated_at"]
RECORD_HEADERS = ["timestamp", "category", "name", "calories", "meal_id"]
LEGACY_RECORD_HEADERS = ["timestamp", "category", "name", "calories"]
CATEGORIES = ["主食", "蛋白質", "蔬菜", "水果", "飲料", "點心", "其他"]
TAIPEI_TIMEZONE = timezone(timedelta(hours=8))


def utc_now():
    return datetime.now(timezone.utc)


def taipei_now():
    return datetime.now(TAIPEI_TIMEZONE)


def main_menu_quick_reply():
    return QuickReply(
        items=[
            QuickReplyButton(action=MessageAction(label="今日進度", text="今日")),
            QuickReplyButton(action=MessageAction(label="近 7 日", text="近7日")),
            QuickReplyButton(action=MessageAction(label="手動新增", text="新增")),
            QuickReplyButton(action=MessageAction(label="飲食比例", text="飲食比例")),
            QuickReplyButton(action=MessageAction(label="查詢熱量", text="查詢大卡")),
            QuickReplyButton(
                action=MessageAction(label="刪除上一筆", text="刪除上一筆")
            ),
            QuickReplyButton(action=MessageAction(label="個人設定", text="設定")),
            QuickReplyButton(action=MessageAction(label="清除全部", text="清除")),
        ]
    )


def menu_message(text):
    return TextSendMessage(text=text, quick_reply=main_menu_quick_reply())


def cancel_quick_reply():
    return QuickReply(
        items=[QuickReplyButton(action=MessageAction(label="取消", text="取消"))]
    )


def category_quick_reply():
    return QuickReply(
        items=[
            QuickReplyButton(action=MessageAction(label=value, text=value))
            for value in CATEGORIES
        ]
        + [QuickReplyButton(action=MessageAction(label="取消", text="取消"))]
    )


def target_quick_reply():
    return QuickReply(
        items=[
            QuickReplyButton(
                action=MessageAction(label="1500 大卡", text="設定目標 1500")
            ),
            QuickReplyButton(
                action=MessageAction(label="1800 大卡", text="設定目標 1800")
            ),
            QuickReplyButton(
                action=MessageAction(label="2000 大卡", text="設定目標 2000")
            ),
            QuickReplyButton(
                action=MessageAction(label="自行輸入", text="自行輸入目標")
            ),
            QuickReplyButton(action=MessageAction(label="暫不設定", text="略過目標")),
        ]
    )


def get_user_target(user_id):
    return user_targets.get(user_id, DAILY_CALORIE_TARGET)


def format_today_progress(user_id, total_calories):
    target = get_user_target(user_id)
    if target <= 0:
        return f"今日已記錄 {total_calories} 大卡"

    difference = target - total_calories
    if difference >= 0:
        return (
            f"今日進度：{total_calories} / {target} 大卡\n"
            f"距離目標還有 {difference} 大卡"
        )
    return (
        f"今日進度：{total_calories} / {target} 大卡\n已超過目標 {abs(difference)} 大卡"
    )


def new_action_token():
    return secrets.token_urlsafe(8)


def pending_action_is_valid(pending, token):
    if not pending or pending.get("token") != token:
        return False
    created_at = pending.get("created_at")
    if not created_at:
        return False
    return utc_now() - created_at < timedelta(minutes=PENDING_ACTION_TTL_MINUTES)


def onboarding_message(prefix=None):
    service_email = service_account_info.get("client_email", "服務帳戶 Email")
    lines = []
    if prefix:
        lines.append(prefix)
    lines.extend(
        [
            "第一次使用需要連結一份 Google Sheet：",
            "1. 建立一份空白試算表",
            f"2. 將編輯權限分享給 {service_email}",
            "3. 回到這裡貼上試算表網址",
        ]
    )
    return "\n".join(lines)


def get_config_sheet():
    global config_sheet

    if not CONFIG_SPREADSHEET_ID:
        raise ValueError("環境變數中找不到 CONFIG_SPREADSHEET_ID")

    if config_sheet is None:
        config_sheet = gc.open_by_key(CONFIG_SPREADSHEET_ID).sheet1

    headers = config_sheet.row_values(1)
    if not headers:
        config_sheet.append_row(CONFIG_HEADERS)
    elif headers == LEGACY_CONFIG_HEADERS:
        config_sheet.update(range_name="A1:D1", values=[CONFIG_HEADERS])
    elif headers != CONFIG_HEADERS:
        raise ValueError("Config Sheet 標題列格式不正確")

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
                if len(row) >= 4 and row[3]:
                    try:
                        user_targets[row[0]] = int(row[3])
                    except ValueError:
                        app.logger.warning("忽略無效的每日目標: %s", row[3])
        app.logger.info("已載入 %s 筆使用者 Sheet 綁定", len(user_sheets))
    except Exception as e:
        app.logger.error("無法載入使用者 Sheet 綁定: %s", e)


def save_user_sheet_binding(user_id, spreadsheet_id):
    sheet = get_config_sheet()
    rows = sheet.get_all_values()
    updated_at = taipei_now().strftime("%Y-%m-%d %H:%M:%S")
    target = user_targets.get(user_id, "")
    values = [user_id, spreadsheet_id, updated_at, target]

    for row_number, row in enumerate(rows[1:], start=2):
        if row and row[0] == user_id:
            sheet.update(
                values=[values],
                range_name=f"A{row_number}:D{row_number}",
            )
            return

    sheet.append_row(values)


def save_user_target(user_id, target):
    if user_id not in user_sheets or not user_sheets[user_id]:
        raise ValueError("使用者尚未連結 Google Sheet")
    previous = user_targets.get(user_id)
    user_targets[user_id] = target
    try:
        save_user_sheet_binding(user_id, user_sheets[user_id])
    except Exception:
        if previous is None:
            user_targets.pop(user_id, None)
        else:
            user_targets[user_id] = previous
        raise


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


def image_confirmation_messages(user_id):
    pending = pending_image_records[user_id]
    detections = pending["detections"]
    token = pending["token"]
    total = sum(int(detection["calories"]) for detection in detections)
    confidence_labels = {"low": "低", "medium": "中", "high": "高"}
    lines = ["照片辨識結果："]
    for index, detection in enumerate(detections, start=1):
        confidence = confidence_labels.get(detection.get("confidence"), "未知")
        lines.append(
            f"{index}. {detection['item']}：{detection['calories']} 大卡"
            f"（可信度 {confidence}）"
        )
    lines.append(f"總計約 {total} 大卡")

    return [
        TextSendMessage(text="\n".join(lines)),
        TemplateSendMessage(
            alt_text=f"辨識到 {len(detections)} 個品項，共約 {total} 大卡",
            template=ButtonsTemplate(
                title="要記錄這次飲食嗎？",
                text=f"共 {len(detections)} 個品項，約 {total} 大卡",
                actions=[
                    PostbackAction(
                        label="確認記錄",
                        data=f"action=confirm_image_record&token={token}",
                        display_text="確認記錄這次飲食",
                    ),
                    PostbackAction(
                        label="修改熱量",
                        data=f"action=edit_image_record&token={token}",
                        display_text="修改辨識熱量",
                    ),
                    PostbackAction(
                        label="重新拍照",
                        data=f"action=retry_image_record&token={token}",
                        display_text="重新拍照",
                    ),
                    PostbackAction(
                        label="取消",
                        data=f"action=cancel_image_record&token={token}",
                        display_text="取消這次記錄",
                    ),
                ],
            ),
        ),
    ]


@handler.add(MessageEvent, message=ImageMessage)
def handle_image_message(event):
    user_id = event.source.user_id
    message_id = event.message.id
    temp_image_path = f"{message_id}.jpg"

    if not user_sheets.get(user_id):
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(
                text=onboarding_message("請先完成設定，再重新傳送這張照片。")
            ),
        )
        return

    try:
        try:
            line_bot_api.push_message(
                user_id,
                TextSendMessage(text="收到照片，正在分析食物與熱量…"),
            )
        except Exception as e:
            app.logger.warning("無法傳送分析中提示: %s", e)

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
            pending_image_records[user_id] = {
                "token": message_id,
                "detections": detections,
                "created_at": utc_now(),
            }
            line_bot_api.reply_message(
                event.reply_token,
                image_confirmation_messages(user_id),
            )
            return
        else:
            response_message = "抱歉，無法辨識圖片中的食物或飲料。"

        line_bot_api.reply_message(event.reply_token, menu_message(response_message))

    except Exception as e:
        app.logger.error(f"Error processing image: {e}")
        line_bot_api.reply_message(
            event.reply_token,
            menu_message("處理圖片時發生錯誤，請稍後重新傳送。"),
        )
    finally:
        if os.path.exists(temp_image_path):
            os.remove(temp_image_path)


@handler.add(PostbackEvent)
def handle_postback(event):
    user_id = event.source.user_id
    params = dict(
        part.split("=", 1) for part in event.postback.data.split("&") if "=" in part
    )
    action = params.get("action")

    if action == "confirm_image_record":
        pending = pending_image_records.get(user_id)
        if not pending_action_is_valid(pending, params.get("token")):
            line_bot_api.reply_message(
                event.reply_token,
                menu_message("這筆照片辨識結果已失效，請重新傳送照片。"),
            )
            return
        detections = pending["detections"]

        try:
            append_image_detections(user_sheets[user_id], detections)
            total = get_today_total(user_sheets[user_id])
        except Exception as e:
            app.logger.error("無法儲存照片辨識結果: %s", e)
            line_bot_api.reply_message(
                event.reply_token,
                menu_message("記錄失敗，請稍後再按一次確認。"),
            )
            return

        pending_image_records.pop(user_id, None)
        line_bot_api.reply_message(
            event.reply_token,
            menu_message(f"已記錄這次飲食。\n{format_today_progress(user_id, total)}"),
        )
        return

    if action == "edit_image_record":
        pending = pending_image_records.get(user_id)
        if not pending_action_is_valid(pending, params.get("token")):
            line_bot_api.reply_message(
                event.reply_token,
                menu_message("這筆照片辨識結果已失效，請重新傳送照片。"),
            )
            return

        detections = pending["detections"]
        user_input_stage[user_id] = "edit_image_calories"
        user_input_data[user_id] = {"token": pending["token"]}
        if len(detections) == 1:
            prompt = (
                f"目前「{detections[0]['item']}」是 {detections[0]['calories']} 大卡。\n"
                "請直接輸入新的熱量，例如：350"
            )
        else:
            prompt = "請輸入「品項編號 新熱量」，例如：2 350"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=prompt, quick_reply=cancel_quick_reply()),
        )
        return

    if action == "retry_image_record":
        pending = pending_image_records.get(user_id)
        if pending_action_is_valid(pending, params.get("token")):
            pending_image_records.pop(user_id, None)
        user_input_stage.pop(user_id, None)
        user_input_data.pop(user_id, None)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(
                text="請重新拍攝或選擇一張較清楚的照片。",
                quick_reply=QuickReply(
                    items=[
                        QuickReplyButton(action=CameraAction(label="開啟相機")),
                        QuickReplyButton(
                            action=MessageAction(label="取消", text="取消")
                        ),
                    ]
                ),
            ),
        )
        return

    if action == "cancel_image_record":
        pending = pending_image_records.get(user_id)
        if pending_action_is_valid(pending, params.get("token")):
            pending_image_records.pop(user_id, None)
            response_message = "已取消這次記錄。"
        else:
            response_message = "這筆照片辨識結果已失效。"
        line_bot_api.reply_message(event.reply_token, menu_message(response_message))
        return

    if action == "confirm_clear":
        pending = pending_clear_actions.get(user_id)
        if not pending_action_is_valid(pending, params.get("token")):
            line_bot_api.reply_message(
                event.reply_token,
                menu_message("這個清除確認已失效，資料沒有變更。"),
            )
            return
        try:
            clear_sheet(user_sheets[user_id])
            pending_image_records.pop(user_id, None)
            pending_clear_actions.pop(user_id, None)
            pending_delete_actions.pop(user_id, None)
            user_input_stage.pop(user_id, None)
            user_input_data.pop(user_id, None)
            response_message = "已清除所有飲食紀錄。"
        except Exception as e:
            response_message = f"無法清除資料: {str(e)}"
        line_bot_api.reply_message(event.reply_token, menu_message(response_message))
        return

    if action == "cancel_clear":
        pending = pending_clear_actions.get(user_id)
        if pending_action_is_valid(pending, params.get("token")):
            pending_clear_actions.pop(user_id, None)
        line_bot_api.reply_message(event.reply_token, menu_message("已取消清除。"))
        return

    if action == "confirm_delete_meal":
        pending = pending_delete_actions.get(user_id)
        if not pending_action_is_valid(pending, params.get("token")):
            line_bot_api.reply_message(
                event.reply_token,
                menu_message("這個刪除確認已失效，資料沒有變更。"),
            )
            return
        try:
            deleted_count = delete_last_meal(
                user_sheets[user_id], pending["meal_id"], pending["row_number"]
            )
            total = get_today_total(user_sheets[user_id])
            pending_delete_actions.pop(user_id, None)
            response_message = (
                f"已撤銷上一餐，共刪除 {deleted_count} 個品項。\n"
                f"{format_today_progress(user_id, total)}"
            )
        except Exception as e:
            response_message = f"無法刪除上一餐: {str(e)}"
        line_bot_api.reply_message(event.reply_token, menu_message(response_message))
        return

    if action == "cancel_delete_meal":
        pending = pending_delete_actions.get(user_id)
        if pending_action_is_valid(pending, params.get("token")):
            pending_delete_actions.pop(user_id, None)
        line_bot_api.reply_message(event.reply_token, menu_message("已取消刪除。"))
        return

    line_bot_api.reply_message(
        event.reply_token, menu_message("這個操作已失效，請重新選擇。")
    )


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
                TextSendMessage(text=onboarding_message()),
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
        user_input_stage[user_id] = "onboarding_target"
        user_input_data[user_id] = {}
        line_bot_api.reply_message(
            event.reply_token,
            [
                TextSendMessage(text="已成功連結到您的 Google Sheet。"),
                TextSendMessage(
                    text="要設定每日熱量目標嗎？之後也能從「個人設定」修改。",
                    quick_reply=target_quick_reply(),
                ),
            ],
        )
        return

    if user_message == "取消":
        user_input_stage.pop(user_id, None)
        user_input_data.pop(user_id, None)
        pending_image_records.pop(user_id, None)
        pending_clear_actions.pop(user_id, None)
        pending_delete_actions.pop(user_id, None)
        line_bot_api.reply_message(event.reply_token, menu_message("已取消操作。"))
        return

    if user_message.startswith("設定目標 "):
        try:
            target = int(user_message.split(" ", 1)[1])
            if target < 500 or target > 10000:
                raise ValueError
            save_user_target(user_id, target)
            response_message = f"每日熱量目標已設定為 {target} 大卡。"
        except ValueError:
            response_message = "每日目標請輸入 500 到 10000 之間的整數。"
        except Exception as e:
            response_message = f"無法儲存每日目標: {str(e)}"
        user_input_stage.pop(user_id, None)
        user_input_data.pop(user_id, None)
        line_bot_api.reply_message(event.reply_token, menu_message(response_message))
        return

    if user_message == "略過目標":
        try:
            save_user_target(user_id, 0)
            response_message = "已略過每日目標；仍會顯示每日累積熱量。"
        except Exception as e:
            response_message = f"無法儲存設定: {str(e)}"
        user_input_stage.pop(user_id, None)
        user_input_data.pop(user_id, None)
        line_bot_api.reply_message(event.reply_token, menu_message(response_message))
        return

    if user_message == "自行輸入目標":
        user_input_stage[user_id] = "target_custom"
        user_input_data[user_id] = {}
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(
                text="請輸入每日目標大卡，例如：1800",
                quick_reply=cancel_quick_reply(),
            ),
        )
        return

    if user_message == "設定":
        target = get_user_target(user_id)
        current = (
            f"目前每日目標是 {target} 大卡。" if target else "目前未設定每日目標。"
        )
        user_input_stage[user_id] = "target_setting"
        user_input_data[user_id] = {}
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(
                text=f"{current}\n請選擇新目標：",
                quick_reply=target_quick_reply(),
            ),
        )
        return

    if user_message in {
        "新增",
        "拍照說明",
        "清除",
        "刪除上一筆",
        "今日",
        "近7日",
        "加總",
        "飲食比例",
        "查詢大卡",
    }:
        user_input_stage.pop(user_id, None)
        user_input_data.pop(user_id, None)

    if user_message == "新增":
        user_input_stage[user_id] = "category"
        user_input_data[user_id] = {}
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(
                text="請選擇食物種類",
                quick_reply=category_quick_reply(),
            ),
        )
        return

    if user_message == "拍照說明":
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(
                text="請拍攝整份餐點，盡量保持光線充足、避免遮擋。",
                quick_reply=QuickReply(
                    items=[
                        QuickReplyButton(action=CameraAction(label="開啟相機")),
                        QuickReplyButton(
                            action=MessageAction(label="取消", text="取消")
                        ),
                    ]
                ),
            ),
        )
        return

    if user_message == "清除":
        token = new_action_token()
        pending_clear_actions[user_id] = {
            "token": token,
            "created_at": utc_now(),
        }
        line_bot_api.reply_message(
            event.reply_token,
            TemplateSendMessage(
                alt_text="確認是否清除所有飲食紀錄",
                template=ButtonsTemplate(
                    title="清除全部紀錄？",
                    text="此操作無法復原，確定要繼續嗎？",
                    actions=[
                        PostbackAction(
                            label="確認清除",
                            data=f"action=confirm_clear&token={token}",
                            display_text="確認清除所有紀錄",
                        ),
                        PostbackAction(
                            label="取消",
                            data=f"action=cancel_clear&token={token}",
                            display_text="取消清除",
                        ),
                    ],
                ),
            ),
        )
        return

    if user_message == "刪除上一筆":
        try:
            summary = get_last_meal_summary(user_sheets[user_id])
        except Exception as e:
            line_bot_api.reply_message(
                event.reply_token, menu_message(f"無法取得上一餐: {str(e)}")
            )
            return

        if not summary:
            line_bot_api.reply_message(
                event.reply_token, menu_message("目前沒有可以刪除的飲食紀錄。")
            )
            return

        token = new_action_token()
        pending_delete_actions[user_id] = {
            "token": token,
            "created_at": utc_now(),
            "meal_id": summary["meal_id"],
            "row_number": summary["row_number"],
        }
        item_text = "、".join(summary["items"])
        button_item_text = item_text if len(item_text) <= 28 else f"{item_text[:27]}…"
        line_bot_api.reply_message(
            event.reply_token,
            TemplateSendMessage(
                alt_text=f"確認是否撤銷上一餐：{item_text}",
                template=ButtonsTemplate(
                    title="撤銷上一餐？",
                    text=f"{button_item_text}，共 {summary['calories']} 大卡",
                    actions=[
                        PostbackAction(
                            label="確認撤銷",
                            data=f"action=confirm_delete_meal&token={token}",
                            display_text="確認撤銷上一餐",
                        ),
                        PostbackAction(
                            label="取消",
                            data=f"action=cancel_delete_meal&token={token}",
                            display_text="取消撤銷",
                        ),
                    ],
                ),
            ),
        )
        return

    if user_message == "今日":
        try:
            total = get_today_total(user_sheets[user_id])
            response_message = format_today_progress(user_id, total)
        except Exception as e:
            response_message = f"無法取得今日進度: {str(e)}"
        line_bot_api.reply_message(event.reply_token, menu_message(response_message))
        return

    if user_message == "近7日":
        try:
            total = sum_calories(user_sheets[user_id], 7)
            response_message = f"近 7 日共記錄 {total} 大卡"
        except Exception as e:
            response_message = f"無法取得近 7 日統計: {str(e)}"
        line_bot_api.reply_message(event.reply_token, menu_message(response_message))
        return

    if user_message == "加總":
        user_input_stage[user_id] = "sum_period"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(
                text="請選擇統計期間",
                quick_reply=QuickReply(
                    items=[
                        QuickReplyButton(
                            action=MessageAction(label="1 天", text="1天")
                        ),
                        QuickReplyButton(
                            action=MessageAction(label="7 天", text="7天")
                        ),
                        QuickReplyButton(
                            action=MessageAction(label="取消", text="取消")
                        ),
                    ]
                ),
            ),
        )
        return

    if user_message == "飲食比例":
        user_input_stage[user_id] = "ratio_period"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(
                text="想查看哪個期間的飲食比例？",
                quick_reply=QuickReply(
                    items=[
                        QuickReplyButton(
                            action=MessageAction(label="今日", text="今日比例")
                        ),
                        QuickReplyButton(
                            action=MessageAction(label="近 7 日", text="近7日比例")
                        ),
                        QuickReplyButton(
                            action=MessageAction(label="全部", text="全部比例")
                        ),
                        QuickReplyButton(
                            action=MessageAction(label="取消", text="取消")
                        ),
                    ]
                ),
            ),
        )
        return

    if user_message == "查詢大卡":
        user_input_stage[user_id] = "query_kcal"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(
                text="請輸入要查詢的食物名稱",
                quick_reply=cancel_quick_reply(),
            ),
        )
        return

    stage = user_input_stage.get(user_id)

    if stage in {"onboarding_target", "target_setting", "target_custom"}:
        try:
            target = int(user_message)
            if target < 500 or target > 10000:
                raise ValueError
            save_user_target(user_id, target)
            response_message = f"每日熱量目標已設定為 {target} 大卡。"
            user_input_stage.pop(user_id, None)
            user_input_data.pop(user_id, None)
            line_bot_api.reply_message(
                event.reply_token, menu_message(response_message)
            )
        except ValueError:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(
                    text="請輸入 500 到 10000 之間的整數。",
                    quick_reply=cancel_quick_reply(),
                ),
            )
        except Exception as e:
            line_bot_api.reply_message(
                event.reply_token, menu_message(f"無法儲存每日目標: {str(e)}")
            )
        return

    if stage == "edit_image_calories":
        pending = pending_image_records.get(user_id)
        token = user_input_data.get(user_id, {}).get("token")
        if not pending_action_is_valid(pending, token):
            user_input_stage.pop(user_id, None)
            user_input_data.pop(user_id, None)
            line_bot_api.reply_message(
                event.reply_token,
                menu_message("這筆辨識結果已失效，請重新傳送照片。"),
            )
            return

        detections = pending["detections"]
        try:
            if len(detections) == 1:
                item_index = 0
                calories = int(user_message)
            else:
                item_text, calorie_text = user_message.split(maxsplit=1)
                item_index = int(item_text) - 1
                calories = int(calorie_text)
            if not 0 <= item_index < len(detections) or calories < 0:
                raise ValueError
        except ValueError:
            example = "350" if len(detections) == 1 else "2 350"
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(
                    text=f"格式不正確，請照範例輸入：{example}",
                    quick_reply=cancel_quick_reply(),
                ),
            )
            return

        detections[item_index]["calories"] = calories
        pending["created_at"] = utc_now()
        user_input_stage.pop(user_id, None)
        user_input_data.pop(user_id, None)
        line_bot_api.reply_message(
            event.reply_token, image_confirmation_messages(user_id)
        )
        return

    if stage == "ratio_period":
        period_options = {
            "今日比例": {"today_only": True, "days": None, "label": "今日"},
            "近7日比例": {"today_only": False, "days": 7, "label": "近 7 日"},
            "全部比例": {"today_only": False, "days": None, "label": "全部"},
        }
        option = period_options.get(user_message)
        if not option:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="請從按鈕選擇今日、近 7 日或全部。"),
            )
            return
        try:
            category_ratios = calculate_category_ratios(
                user_sheets[user_id],
                days=option["days"],
                today_only=option["today_only"],
            )
            if category_ratios:
                response_message = f"{option['label']}飲食比例：\n" + "\n".join(
                    f"{category}: {ratio:.1f}%"
                    for category, ratio in category_ratios.items()
                )
            else:
                response_message = f"{option['label']}還沒有可統計的飲食紀錄。"
        except Exception as e:
            response_message = f"無法計算飲食比例: {str(e)}"
        user_input_stage.pop(user_id, None)
        line_bot_api.reply_message(event.reply_token, menu_message(response_message))
        return

    if stage == "query_kcal":
        result = search_kcal(user_message)
        response_message = (
            result or f"抱歉，我們的資料庫中沒有 '{user_message}' 的熱量資訊"
        )
        user_input_stage[user_id] = None
        line_bot_api.reply_message(event.reply_token, menu_message(response_message))
        return

    if stage == "category":
        if user_message not in CATEGORIES:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(
                    text="請從按鈕選擇食物種類。",
                    quick_reply=category_quick_reply(),
                ),
            )
            return
        user_input_data.setdefault(user_id, {})["category"] = user_message
        user_input_stage[user_id] = "name"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="請輸入名稱", quick_reply=cancel_quick_reply()),
        )
        return

    if stage == "name":
        user_input_data.setdefault(user_id, {})["name"] = user_message
        user_input_stage[user_id] = "calories"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="請輸入大卡", quick_reply=cancel_quick_reply()),
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
                TextSendMessage(
                    text="大卡必須是大於或等於 0 的整數，請重新輸入",
                    quick_reply=cancel_quick_reply(),
                ),
            )
            return

        user_input_data.setdefault(user_id, {})["calories"] = calories
        user_input_data[user_id]["timestamp"] = taipei_now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        try:
            append_values(user_sheets[user_id], user_input_data[user_id])
            total = get_today_total(user_sheets[user_id])
            response_message = (
                f"已儲存到 Google Sheets。\n{format_today_progress(user_id, total)}"
            )
        except Exception as e:
            response_message = f"無法儲存訊息到 Google Sheets: {str(e)}"

        user_input_stage[user_id] = None
        user_input_data[user_id] = {}
        line_bot_api.reply_message(event.reply_token, menu_message(response_message))
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
        line_bot_api.reply_message(event.reply_token, menu_message(response_message))
        return

    line_bot_api.reply_message(
        event.reply_token,
        menu_message("我不確定你想做什麼，請從下方選擇功能，或直接傳送食物照片。"),
    )


def search_kcal(name):
    for key, values in food_dict.items():
        if name in key:
            if not values:
                continue
            calories = values[0]
            serving = values[1] if len(values) > 1 and values[1] else "資料庫標示份量"
            return f"品項：{key}, 熱量：{calories} kcal, 份量：{serving}"
    return None


def get_record_sheet(spreadsheet_id):
    sheet = gc.open_by_key(spreadsheet_id).sheet1
    headers = sheet.row_values(1)
    if not headers:
        sheet.append_row(RECORD_HEADERS)
    elif headers == LEGACY_RECORD_HEADERS:
        sheet.update(range_name="A1:E1", values=[RECORD_HEADERS])
    elif headers != RECORD_HEADERS:
        raise ValueError("飲食 Sheet 標題列格式不正確")
    return sheet


def add_headers(spreadsheet_id):
    get_record_sheet(spreadsheet_id)


def append_values(spreadsheet_id, data):
    sheet = get_record_sheet(spreadsheet_id)
    row = [
        data["timestamp"],
        data["category"],
        data["name"],
        data["calories"],
        uuid.uuid4().hex,
    ]
    sheet.append_row(row)


def append_image_detections(spreadsheet_id, detections):
    sheet = get_record_sheet(spreadsheet_id)
    meal_id = uuid.uuid4().hex
    timestamp = taipei_now().strftime("%Y-%m-%d %H:%M:%S")
    rows = [
        [
            timestamp,
            detection.get("category", "圖片辨識"),
            detection["item"],
            int(detection["calories"]),
            meal_id,
        ]
        for detection in detections
    ]
    sheet.append_rows(rows, value_input_option="USER_ENTERED")
    return meal_id


def get_today_total(spreadsheet_id):
    sheet = get_record_sheet(spreadsheet_id)
    all_records = sheet.get_all_records()
    today = taipei_now().date()
    total = 0

    for record in all_records:
        try:
            timestamp = datetime.strptime(str(record["timestamp"]), "%Y-%m-%d %H:%M:%S")
            if timestamp.date() == today:
                total += int(record["calories"])
        except (KeyError, TypeError, ValueError):
            app.logger.warning("略過格式不正確的飲食紀錄: %s", record)

    return total


def clear_sheet(spreadsheet_id):
    sheet = get_record_sheet(spreadsheet_id)
    sheet.clear()
    sheet.append_row(RECORD_HEADERS)


def get_last_meal_summary(spreadsheet_id):
    sheet = get_record_sheet(spreadsheet_id)
    values = sheet.get_all_values()
    records = []
    for row_number, row in enumerate(values[1:], start=2):
        padded = row + [""] * (5 - len(row))
        if any(padded[:4]):
            records.append((row_number, padded))

    if not records:
        return None

    last_row_number, last_row = records[-1]
    meal_id = last_row[4]
    if meal_id:
        meal_records = [record for record in records if record[1][4] == meal_id]
    else:
        meal_records = [(last_row_number, last_row)]

    items = [row[2] or "未命名品項" for _, row in meal_records]
    calories = 0
    for _, row in meal_records:
        try:
            calories += int(row[3])
        except ValueError:
            pass

    return {
        "meal_id": meal_id,
        "row_number": last_row_number,
        "items": items,
        "calories": calories,
    }


def delete_last_meal(spreadsheet_id, meal_id, legacy_row_number):
    sheet = get_record_sheet(spreadsheet_id)
    values = sheet.get_all_values()
    if meal_id:
        row_numbers = [
            row_number
            for row_number, row in enumerate(values[1:], start=2)
            if len(row) >= 5 and row[4] == meal_id
        ]
    elif 1 < legacy_row_number <= len(values):
        row_numbers = [legacy_row_number]
    else:
        row_numbers = []

    if not row_numbers:
        raise ValueError("找不到要撤銷的上一餐，可能已被變更")

    for row_number in reversed(row_numbers):
        sheet.delete_rows(row_number)
    return len(row_numbers)


def sum_calories(spreadsheet_id, days):
    sheet = get_record_sheet(spreadsheet_id)
    all_records = sheet.get_all_records()
    now = taipei_now()
    total_calories = 0
    for record in all_records:
        try:
            timestamp = datetime.strptime(
                str(record["timestamp"]), "%Y-%m-%d %H:%M:%S"
            ).replace(tzinfo=TAIPEI_TIMEZONE)
            if timedelta(0) <= now - timestamp < timedelta(days=days):
                total_calories += int(record["calories"])
        except (KeyError, TypeError, ValueError):
            app.logger.warning("略過格式不正確的飲食紀錄: %s", record)
    return total_calories


def calculate_category_ratios(spreadsheet_id, days=None, today_only=False):
    sheet = get_record_sheet(spreadsheet_id)
    all_records = sheet.get_all_records()
    category_totals = {}
    total_calories = 0
    now = taipei_now()

    for record in all_records:
        try:
            timestamp = datetime.strptime(
                str(record["timestamp"]), "%Y-%m-%d %H:%M:%S"
            ).replace(tzinfo=TAIPEI_TIMEZONE)
            if today_only and timestamp.date() != now.date():
                continue
            if days is not None and not timedelta(0) <= now - timestamp < timedelta(
                days=days
            ):
                continue
            category = str(record["category"]) or "其他"
            calories = int(record["calories"])
        except (KeyError, TypeError, ValueError):
            app.logger.warning("略過格式不正確的飲食紀錄: %s", record)
            continue
        total_calories += calories
        if category in category_totals:
            category_totals[category] += calories
        else:
            category_totals[category] = calories

    if total_calories == 0:
        return {}

    category_ratios = {
        category: (calories / total_calories) * 100
        for category, calories in category_totals.items()
    }
    return category_ratios


if __name__ == "__main__":
    app.run(debug=True)
