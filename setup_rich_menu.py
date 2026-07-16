import argparse
import io
import os

from dotenv import load_dotenv
from linebot import LineBotApi
from linebot.models import (
    MessageAction,
    RichMenu,
    RichMenuArea,
    RichMenuBounds,
    RichMenuSize,
)
from PIL import Image, ImageDraw, ImageFont


MENU_NAME = "drinkControl Main Menu v1"
MENU_WIDTH = 2500
MENU_HEIGHT = 1686
MENU_ITEMS = [
    ("拍照記錄", "拍照說明", "CAMERA"),
    ("今日進度", "今日", "TODAY"),
    ("近 7 日", "近7日", "7 DAYS"),
    ("手動新增", "新增", "ADD"),
    ("飲食比例", "飲食比例", "RATIO"),
    ("個人設定", "設定", "SETTINGS"),
]
COLORS = ["#E9F5EE", "#EAF0FA", "#F7F0E5", "#F8EBE8", "#EFEAFA", "#E8F4F6"]


def load_font(size):
    candidates = [
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


def build_menu_image():
    image = Image.new("RGB", (MENU_WIDTH, MENU_HEIGHT), "#FAFAF7")
    draw = ImageDraw.Draw(image)
    title_font = load_font(92)
    subtitle_font = load_font(40)
    cell_width = MENU_WIDTH // 3
    cell_height = MENU_HEIGHT // 2

    for index, (title, _message, subtitle) in enumerate(MENU_ITEMS):
        column = index % 3
        row = index // 3
        left = column * cell_width
        top = row * cell_height
        right = MENU_WIDTH if column == 2 else left + cell_width
        bottom = MENU_HEIGHT if row == 1 else top + cell_height
        draw.rectangle((left, top, right, bottom), fill=COLORS[index])
        draw.line((right - 1, top, right - 1, bottom), fill="#D6D8D4", width=3)
        draw.line((left, bottom - 1, right, bottom - 1), fill="#D6D8D4", width=3)

        title_box = draw.textbbox((0, 0), title, font=title_font)
        title_width = title_box[2] - title_box[0]
        title_height = title_box[3] - title_box[1]
        draw.text(
            (left + (right - left - title_width) / 2, top + cell_height * 0.39),
            title,
            fill="#24332C",
            font=title_font,
        )
        subtitle_box = draw.textbbox((0, 0), subtitle, font=subtitle_font)
        subtitle_width = subtitle_box[2] - subtitle_box[0]
        draw.text(
            (
                left + (right - left - subtitle_width) / 2,
                top + cell_height * 0.39 + title_height + 42,
            ),
            subtitle,
            fill="#66736C",
            font=subtitle_font,
        )

    buffer = io.BytesIO()
    image.save(buffer, format="PNG", optimize=True)
    return buffer.getvalue()


def build_rich_menu():
    cell_width = MENU_WIDTH // 3
    cell_height = MENU_HEIGHT // 2
    areas = []
    for index, (_title, message, _subtitle) in enumerate(MENU_ITEMS):
        column = index % 3
        row = index // 3
        width = MENU_WIDTH - column * cell_width if column == 2 else cell_width
        height = MENU_HEIGHT - row * cell_height if row == 1 else cell_height
        areas.append(
            RichMenuArea(
                bounds=RichMenuBounds(
                    x=column * cell_width,
                    y=row * cell_height,
                    width=width,
                    height=height,
                ),
                action=MessageAction(label=message, text=message),
            )
        )
    return RichMenu(
        size=RichMenuSize(width=MENU_WIDTH, height=MENU_HEIGHT),
        selected=True,
        name=MENU_NAME,
        chat_bar_text="開啟飲食選單",
        areas=areas,
    )


def install_rich_menu(access_token):
    api = LineBotApi(access_token)
    existing = next(
        (menu for menu in api.get_rich_menu_list() if menu.name == MENU_NAME),
        None,
    )
    if existing:
        api.set_default_rich_menu(existing.rich_menu_id)
        return existing.rich_menu_id, False

    rich_menu_id = api.create_rich_menu(build_rich_menu())
    api.set_rich_menu_image(rich_menu_id, "image/png", build_menu_image())
    api.set_default_rich_menu(rich_menu_id)
    return rich_menu_id, True


def main():
    parser = argparse.ArgumentParser(description="建立 drinkControl LINE Rich Menu")
    parser.add_argument(
        "--preview",
        metavar="PATH",
        help="只產生預覽 PNG，不呼叫 LINE API",
    )
    args = parser.parse_args()

    if args.preview:
        with open(args.preview, "wb") as output:
            output.write(build_menu_image())
        print(f"Rich Menu 預覽已輸出至 {args.preview}")
        return

    load_dotenv()
    access_token = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
    if not access_token:
        raise ValueError("環境變數中找不到 LINE_CHANNEL_ACCESS_TOKEN")
    rich_menu_id, created = install_rich_menu(access_token)
    action = "建立並設為預設" if created else "設為預設"
    print(f"Rich Menu 已{action}: {rich_menu_id}")


if __name__ == "__main__":
    main()
