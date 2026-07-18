from linebot.models import FlexSendMessage


def _postback_button(label, data, color="#2F6B52"):
    return {
        "type": "button",
        "height": "sm",
        "style": "link",
        "color": color,
        "action": {
            "type": "postback",
            "label": label,
            "data": data,
            "displayText": label,
        },
    }


def meal_bubble(meal):
    timestamp = meal["timestamp"]
    item_contents = []
    visible_items = meal["items"][:8]
    for item in visible_items:
        item_contents.append(
            {
                "type": "box",
                "layout": "horizontal",
                "contents": [
                    {
                        "type": "text",
                        "text": item["name"],
                        "size": "sm",
                        "color": "#38433D",
                        "flex": 4,
                        "wrap": True,
                    },
                    {
                        "type": "text",
                        "text": f"{item['calories']} kcal",
                        "size": "sm",
                        "color": "#68736D",
                        "align": "end",
                        "flex": 2,
                    },
                ],
                "spacing": "sm",
            }
        )
    if len(meal["items"]) > len(visible_items):
        item_contents.append(
            {
                "type": "text",
                "text": f"另有 {len(meal['items']) - len(visible_items)} 個品項",
                "size": "xs",
                "color": "#8A948E",
            }
        )

    meal_id = meal["meal_id"]
    return {
        "type": "bubble",
        "size": "kilo",
        "header": {
            "type": "box",
            "layout": "vertical",
            "backgroundColor": "#E9F5EE",
            "paddingAll": "16px",
            "contents": [
                {
                    "type": "text",
                    "text": timestamp.strftime("%m/%d %H:%M"),
                    "size": "xs",
                    "color": "#527161",
                    "weight": "bold",
                },
                {
                    "type": "text",
                    "text": meal.get("title") or "飲食紀錄",
                    "size": "lg",
                    "weight": "bold",
                    "color": "#24332C",
                    "margin": "sm",
                    "wrap": True,
                },
            ],
        },
        "body": {
            "type": "box",
            "layout": "vertical",
            "spacing": "md",
            "contents": item_contents
            + [
                {"type": "separator", "margin": "md"},
                {
                    "type": "box",
                    "layout": "horizontal",
                    "margin": "md",
                    "contents": [
                        {
                            "type": "text",
                            "text": "總熱量",
                            "weight": "bold",
                            "color": "#38433D",
                        },
                        {
                            "type": "text",
                            "text": f"{meal['total_calories']} kcal",
                            "weight": "bold",
                            "align": "end",
                            "color": "#2F6B52",
                        },
                    ],
                },
            ],
        },
        "footer": {
            "type": "box",
            "layout": "horizontal",
            "spacing": "xs",
            "paddingAll": "8px",
            "contents": [
                _postback_button("編輯", f"action=edit_meal&meal_id={meal_id}"),
                _postback_button("再記一次", f"action=repeat_meal&meal_id={meal_id}"),
                _postback_button(
                    "刪除", f"action=request_delete_meal&meal_id={meal_id}", "#A65353"
                ),
            ],
        },
    }


def meal_cards_message(meals, title):
    bubbles = [meal_bubble(meal) for meal in meals[:12]]
    contents = (
        bubbles[0] if len(bubbles) == 1 else {"type": "carousel", "contents": bubbles}
    )
    return FlexSendMessage(alt_text=f"{title}：共 {len(meals)} 餐", contents=contents)


def weekly_report_message(report):
    metric_rows = [
        ("本週總熱量", f"{report['total_calories']} kcal"),
        ("每日平均", f"{report['daily_average']} kcal"),
        ("記錄天數", f"{report['recorded_days']} / 7 天"),
    ]
    if report.get("target"):
        metric_rows.append(("目標內天數", f"{report['goal_days']} 天"))

    metrics = []
    for label, value in metric_rows:
        metrics.append(
            {
                "type": "box",
                "layout": "horizontal",
                "contents": [
                    {"type": "text", "text": label, "size": "sm", "color": "#68736D"},
                    {
                        "type": "text",
                        "text": value,
                        "size": "sm",
                        "weight": "bold",
                        "align": "end",
                        "color": "#24332C",
                    },
                ],
            }
        )

    category_text = (
        "、".join(
            f"{name} {ratio:.0f}%" for name, ratio in report["category_ratios"][:3]
        )
        or "尚無資料"
    )
    top_item_text = (
        f"最高熱量品項：{report['top_items'][0][0]}" if report.get("top_items") else ""
    )
    highest_day_text = (
        f"最高的一天：{report['highest_day'].strftime('%m/%d')}・"
        f"{report['highest_total']} kcal"
    )
    comparison = report.get("comparison_text") or "上週尚無足夠資料"
    insight = report.get("insight") or "持續記錄，下一份週報會更有參考價值。"

    bubble = {
        "type": "bubble",
        "header": {
            "type": "box",
            "layout": "vertical",
            "backgroundColor": "#EAF0FA",
            "paddingAll": "18px",
            "contents": [
                {
                    "type": "text",
                    "text": "每週飲食報告",
                    "weight": "bold",
                    "size": "xl",
                    "color": "#26364A",
                },
                {
                    "type": "text",
                    "text": report["date_range"],
                    "size": "xs",
                    "color": "#60738A",
                    "margin": "sm",
                },
            ],
        },
        "body": {
            "type": "box",
            "layout": "vertical",
            "spacing": "md",
            "contents": metrics
            + [
                {"type": "separator", "margin": "md"},
                {
                    "type": "text",
                    "text": comparison,
                    "size": "sm",
                    "color": "#43566B",
                    "wrap": True,
                },
                {
                    "type": "text",
                    "text": f"主要分類：{category_text}",
                    "size": "sm",
                    "color": "#43566B",
                    "wrap": True,
                },
                {
                    "type": "text",
                    "text": highest_day_text,
                    "size": "sm",
                    "color": "#43566B",
                    "wrap": True,
                },
                {
                    "type": "text",
                    "text": top_item_text,
                    "size": "sm",
                    "color": "#43566B",
                    "wrap": True,
                },
                {
                    "type": "box",
                    "layout": "vertical",
                    "backgroundColor": "#F5F7FA",
                    "cornerRadius": "8px",
                    "paddingAll": "12px",
                    "contents": [
                        {
                            "type": "text",
                            "text": insight,
                            "size": "sm",
                            "color": "#38495D",
                            "wrap": True,
                        }
                    ],
                },
            ],
        },
    }
    return FlexSendMessage(
        alt_text=f"每週飲食報告：{report['date_range']}", contents=bubble
    )
