import base64
import json
import os

from openai import OpenAI


FOOD_ANALYSIS_SCHEMA = {
    "type": "object",
    "properties": {
        "detections": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "item": {"type": "string"},
                    "category": {
                        "type": "string",
                        "enum": [
                            "主食",
                            "蛋白質",
                            "蔬菜",
                            "水果",
                            "飲料",
                            "點心",
                            "其他",
                        ],
                    },
                    "calories": {"type": "integer", "minimum": 0},
                    "confidence": {
                        "type": "string",
                        "enum": ["low", "medium", "high"],
                    },
                },
                "required": ["item", "category", "calories", "confidence"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["detections"],
    "additionalProperties": False,
}

ANALYSIS_PROMPT = """分析照片中可食用的食物與飲料。
辨識每個主要品項，以繁體中文命名，並依照片可見份量估算整份熱量（大卡）。
category 請依品項歸類為主食、蛋白質、蔬菜、水果、飲料、點心或其他。
不要把餐具、容器或同一品項的不同部分重複計算。
若無法辨識任何食物或飲料，detections 回傳空陣列。
confidence 表示品項與份量估算的整體可信度。"""


def analyze_image(image_path):
    if os.getenv("CV_PROVIDER", "openai").lower() != "openai":
        raise ValueError("CV_PROVIDER 目前僅支援 openai")

    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("環境變數中找不到 OPENAI_API_KEY")

    with open(image_path, "rb") as image_file:
        image_base64 = base64.b64encode(image_file.read()).decode("ascii")

    client = OpenAI(api_key=api_key, timeout=30.0, max_retries=1)
    response = client.responses.create(
        model=os.getenv("OPENAI_VISION_MODEL", "gpt-4.1-mini"),
        input=[
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": ANALYSIS_PROMPT},
                    {
                        "type": "input_image",
                        "image_url": f"data:image/jpeg;base64,{image_base64}",
                        "detail": os.getenv("OPENAI_IMAGE_DETAIL", "low"),
                    },
                ],
            }
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "food_analysis",
                "strict": True,
                "schema": FOOD_ANALYSIS_SCHEMA,
            }
        },
        max_output_tokens=300,
    )

    result = json.loads(response.output_text)
    return result if result["detections"] else None
