import os
import sys
import unittest
from datetime import timedelta
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ.update(
    {
        "LINE_CHANNEL_ACCESS_TOKEN": "test",
        "LINE_CHANNEL_SECRET": "test",
        "GOOGLE_SERVICE_ACCOUNT_INFO": "{}",
        "CONFIG_SPREADSHEET_ID": "",
        "OPENAI_API_KEY": "test",
    }
)

from google.oauth2.service_account import Credentials  # noqa: E402
import gspread  # noqa: E402


Credentials.from_service_account_info = classmethod(
    lambda cls, info, scopes=None: object()
)
gspread.authorize = lambda credentials: None

import app  # noqa: E402
import setup_rich_menu  # noqa: E402


class FakeSheet:
    def __init__(self, values=None):
        self.values = [list(row) for row in (values or [])]

    def _ensure_size(self, row_count, column_count):
        while len(self.values) < row_count:
            self.values.append([])
        for row in self.values:
            row.extend([""] * (column_count - len(row)))

    def row_values(self, row_number):
        if row_number > len(self.values):
            return []
        row = list(self.values[row_number - 1])
        while row and row[-1] == "":
            row.pop()
        return row

    def append_row(self, row):
        self.values.append(list(row))

    def append_rows(self, rows, value_input_option=None):
        self.values.extend(list(row) for row in rows)

    def update(self, range_name, values):
        start, end = range_name.split(":")
        start_column = ord(start[0]) - ord("A")
        row_number = int(start[1:])
        end_column = ord(end[0]) - ord("A")
        self._ensure_size(row_number, end_column + 1)
        for offset, value in enumerate(values[0]):
            self.values[row_number - 1][start_column + offset] = value

    def get_all_values(self):
        return [list(row) for row in self.values]

    def get_all_records(self):
        if not self.values:
            return []
        headers = self.values[0]
        records = []
        for row in self.values[1:]:
            padded = row + [""] * (len(headers) - len(row))
            if any(padded):
                records.append(dict(zip(headers, padded)))
        return records

    def delete_rows(self, row_number):
        del self.values[row_number - 1]

    def clear(self):
        self.values = []


class FakeBook:
    def __init__(self, sheet):
        self.sheet1 = sheet


class FakeGC:
    def __init__(self, sheets):
        self.sheets = sheets

    def open_by_key(self, key):
        return FakeBook(self.sheets[key])


class FakeBot:
    def __init__(self):
        self.replies = []

    def reply_message(self, token, message):
        self.replies.append((token, message))


class FakeImageContent:
    def iter_content(self):
        return [b"fake-image"]


class FakeImageBot(FakeBot):
    def __init__(self):
        super().__init__()
        self.pushes = []

    def push_message(self, user_id, message):
        self.pushes.append((user_id, message))

    def get_message_content(self, message_id):
        return FakeImageContent()


class Obj:
    pass


def postback_event(data, user_id="u1"):
    event = Obj()
    event.reply_token = "reply"
    event.source = Obj()
    event.source.user_id = user_id
    event.postback = Obj()
    event.postback.data = data
    return event


def text_event(text, user_id="u1"):
    event = Obj()
    event.reply_token = "reply"
    event.source = Obj()
    event.source.user_id = user_id
    event.message = Obj()
    event.message.text = text
    return event


def image_event(message_id="flow-test", user_id="u1"):
    event = Obj()
    event.reply_token = "reply"
    event.source = Obj()
    event.source.user_id = user_id
    event.message = Obj()
    event.message.id = message_id
    return event


class DrinkControlFlowTests(unittest.TestCase):
    def setUp(self):
        self.sheet = FakeSheet([app.LEGACY_RECORD_HEADERS])
        self.config = FakeSheet([app.CONFIG_HEADERS])
        app.gc = FakeGC({"sheet": self.sheet, "config": self.config})
        app.CONFIG_SPREADSHEET_ID = "config"
        app.config_sheet = None
        app.line_bot_api = FakeBot()
        app.user_sheets.clear()
        app.user_targets.clear()
        app.user_input_stage.clear()
        app.user_input_data.clear()
        app.pending_image_records.clear()
        app.pending_clear_actions.clear()
        app.pending_delete_actions.clear()
        app.user_sheets["u1"] = "sheet"

    def test_onboarding_links_sheet_then_sets_personal_target(self):
        app.user_sheets.clear()
        app.handle_message(text_event("你好"))
        self.assertIn("第一次使用", app.line_bot_api.replies[-1][1].text)

        app.handle_message(
            text_event("https://docs.google.com/spreadsheets/d/sheet/edit")
        )
        self.assertEqual(app.user_sheets["u1"], "sheet")
        self.assertEqual(app.user_input_stage["u1"], "onboarding_target")
        target_prompt = app.line_bot_api.replies[-1][1][-1]
        self.assertIsNotNone(target_prompt.quick_reply)

        app.handle_message(text_event("設定目標 1800"))
        self.assertEqual(app.user_targets["u1"], 1800)
        self.assertEqual(self.config.values[1][3], 1800)

    def test_manual_add_uses_category_buttons_and_creates_meal(self):
        app.handle_message(text_event("新增"))
        prompt = app.line_bot_api.replies[-1][1]
        labels = [item.action.label for item in prompt.quick_reply.items]
        self.assertTrue(set(app.CATEGORIES).issubset(labels))

        app.handle_message(text_event("飲料"))
        app.handle_message(text_event("無糖豆漿"))
        app.handle_message(text_event("120"))
        self.assertEqual(self.sheet.values[1][1:4], ["飲料", "無糖豆漿", 120])
        self.assertTrue(self.sheet.values[1][4])

    def test_legacy_sheet_migrates_without_shifting_data(self):
        self.sheet.values.append(["2026-07-17 08:00:00", "主食", "白飯", "280"])
        app.get_record_sheet("sheet")
        self.assertEqual(self.sheet.values[0], app.RECORD_HEADERS)
        self.assertEqual(
            self.sheet.values[1][:4],
            ["2026-07-17 08:00:00", "主食", "白飯", "280"],
        )

    def test_photo_items_share_meal_id_and_delete_as_one_meal(self):
        app.append_image_detections(
            "sheet",
            [
                {"item": "白飯", "category": "主食", "calories": 280},
                {"item": "雞胸肉", "category": "蛋白質", "calories": 180},
            ],
        )
        meal_ids = {row[4] for row in self.sheet.values[1:]}
        self.assertEqual(len(meal_ids), 1)
        summary = app.get_last_meal_summary("sheet")
        self.assertEqual(summary["calories"], 460)
        self.assertEqual(summary["items"], ["白飯", "雞胸肉"])
        deleted = app.delete_last_meal(
            "sheet", summary["meal_id"], summary["row_number"]
        )
        self.assertEqual(deleted, 2)
        self.assertEqual(len(self.sheet.values), 1)

    def test_confirm_delete_postback_removes_whole_meal(self):
        app.append_image_detections(
            "sheet",
            [
                {"item": "白飯", "category": "主食", "calories": 280},
                {"item": "雞胸肉", "category": "蛋白質", "calories": 180},
            ],
        )
        summary = app.get_last_meal_summary("sheet")
        app.pending_delete_actions["u1"] = {
            "token": "delete-1",
            "created_at": app.utc_now(),
            "meal_id": summary["meal_id"],
            "row_number": summary["row_number"],
        }
        app.handle_postback(postback_event("action=confirm_delete_meal&token=delete-1"))
        self.assertEqual(self.sheet.values, [app.RECORD_HEADERS])
        self.assertIn("2 個品項", app.line_bot_api.replies[-1][1].text)

    def test_expired_clear_confirmation_cannot_clear_data(self):
        self.sheet.values.append(["2026-07-17 08:00:00", "主食", "白飯", "280", "m1"])
        app.pending_clear_actions["u1"] = {
            "token": "old",
            "created_at": app.utc_now()
            - timedelta(minutes=app.PENDING_ACTION_TTL_MINUTES + 1),
        }
        app.handle_postback(postback_event("action=confirm_clear&token=old"))
        self.assertEqual(len(self.sheet.values), 2)
        reply = app.line_bot_api.replies[-1][1]
        self.assertIn("失效", reply.text)

    def test_valid_clear_confirmation_is_one_time(self):
        self.sheet.values.append(["2026-07-17 08:00:00", "主食", "白飯", "280", "m1"])
        app.pending_clear_actions["u1"] = {
            "token": "clear-1",
            "created_at": app.utc_now(),
        }
        event = postback_event("action=confirm_clear&token=clear-1")
        app.handle_postback(event)
        self.assertEqual(self.sheet.values, [app.RECORD_HEADERS])
        app.handle_postback(event)
        self.assertEqual(self.sheet.values, [app.RECORD_HEADERS])
        self.assertIn("失效", app.line_bot_api.replies[-1][1].text)

    def test_image_calories_can_be_edited_before_confirmation(self):
        app.pending_image_records["u1"] = {
            "token": "image-1",
            "created_at": app.utc_now(),
            "detections": [
                {
                    "item": "雞肉便當",
                    "category": "主食",
                    "calories": 720,
                    "confidence": "medium",
                }
            ],
        }
        app.user_input_stage["u1"] = "edit_image_calories"
        app.user_input_data["u1"] = {"token": "image-1"}
        app.handle_message(text_event("650"))
        self.assertEqual(
            app.pending_image_records["u1"]["detections"][0]["calories"],
            650,
        )
        messages = app.line_bot_api.replies[-1][1]
        self.assertEqual(len(messages[1].template.actions), 4)

    def test_photo_flow_acknowledges_confirms_and_records(self):
        bot = FakeImageBot()
        app.line_bot_api = bot
        original_compress = app.compress_image
        original_analyze = app.analyze_image
        app.compress_image = lambda *args, **kwargs: None
        app.analyze_image = lambda path: {
            "detections": [
                {
                    "item": "雞肉便當",
                    "category": "主食",
                    "calories": 720,
                    "confidence": "high",
                }
            ]
        }
        try:
            app.handle_image_message(image_event())
        finally:
            app.compress_image = original_compress
            app.analyze_image = original_analyze

        self.assertEqual(len(bot.pushes), 1)
        pending = app.pending_image_records["u1"]
        messages = bot.replies[-1][1]
        self.assertEqual(len(messages[1].template.actions), 4)

        app.handle_postback(
            postback_event(f"action=confirm_image_record&token={pending['token']}")
        )
        self.assertNotIn("u1", app.pending_image_records)
        self.assertEqual(self.sheet.values[1][2:4], ["雞肉便當", 720])
        self.assertIn("今日已記錄 720", bot.replies[-1][1].text)

    def test_user_targets_are_independent(self):
        app.user_targets["u1"] = 1800
        app.user_targets["u2"] = 2200
        self.assertIn("1300", app.format_today_progress("u1", 500))
        self.assertIn("1700", app.format_today_progress("u2", 500))

    def test_target_is_saved_in_migrated_config_sheet(self):
        self.config.values = [
            app.LEGACY_CONFIG_HEADERS,
            ["u1", "sheet", "2026-07-17 08:00:00"],
        ]
        app.config_sheet = None
        app.save_user_target("u1", 1900)
        self.assertEqual(self.config.values[0], app.CONFIG_HEADERS)
        self.assertEqual(self.config.values[1][3], 1900)

    def test_ratio_period_filters_today(self):
        today = app.taipei_now().strftime("%Y-%m-%d %H:%M:%S")
        old = (app.taipei_now() - timedelta(days=10)).strftime("%Y-%m-%d %H:%M:%S")
        self.sheet.values = [
            app.RECORD_HEADERS,
            [today, "主食", "白飯", "300", "m1"],
            [today, "蛋白質", "雞肉", "100", "m1"],
            [old, "點心", "蛋糕", "500", "m2"],
        ]
        ratios = app.calculate_category_ratios("sheet", today_only=True)
        self.assertEqual(ratios, {"主食": 75.0, "蛋白質": 25.0})

        app.handle_message(text_event("飲食比例"))
        app.handle_message(text_event("今日比例"))
        self.assertIn("主食: 75.0%", app.line_bot_api.replies[-1][1].text)

    def test_rich_menu_has_six_full_coverage_areas(self):
        rich_menu = setup_rich_menu.build_rich_menu()
        self.assertEqual(len(rich_menu.areas), 6)
        self.assertEqual(rich_menu.size.width, 2500)
        self.assertEqual(rich_menu.size.height, 1686)
        image = setup_rich_menu.build_menu_image()
        self.assertGreater(len(image), 10_000)

    def test_rich_menu_installer_uploads_and_sets_default(self):
        calls = []

        class FakeRichMenuApi:
            def __init__(self, token):
                calls.append(("token", token))

            def get_rich_menu_list(self):
                return []

            def create_rich_menu(self, rich_menu):
                calls.append(("create", rich_menu.name))
                return "rich-menu-1"

            def set_rich_menu_image(self, rich_menu_id, content_type, content):
                calls.append(("image", rich_menu_id, content_type, len(content)))

            def set_default_rich_menu(self, rich_menu_id):
                calls.append(("default", rich_menu_id))

        original = setup_rich_menu.LineBotApi
        setup_rich_menu.LineBotApi = FakeRichMenuApi
        try:
            rich_menu_id, created = setup_rich_menu.install_rich_menu("token")
        finally:
            setup_rich_menu.LineBotApi = original

        self.assertTrue(created)
        self.assertEqual(rich_menu_id, "rich-menu-1")
        self.assertIn(("default", "rich-menu-1"), calls)
        self.assertEqual(calls[2][2], "image/png")


if __name__ == "__main__":
    unittest.main()
