from ultralytics import YOLO

model = None

FOOD_CALORIES = {
    "bottle": ("瓶裝飲料", 150),
    "wine glass": ("葡萄酒/酒杯", 125),
    "cup": ("杯裝飲料", 100),
    "banana": ("香蕉", 105),
    "apple": ("蘋果", 95),
    "orange": ("橘子", 62),
    "broccoli": ("花椰菜", 55),
    "carrot": ("胡蘿蔔", 25),
    "hot dog": ("熱狗", 290),
    "pizza": ("披薩", 285),
    "donut": ("甜甜圈", 269),
    "cake": ("蛋糕", 350),
    "bowl": ("碗裝食物", 200),
    "sandwich": ("三明治", 350),
}


def analyze_image(image_path):
    global model
    if model is None:
        print("Initializing YOLO model...")
        model = YOLO("yolov8n.pt")
        print("YOLO model initialized.")

    results = model.predict(image_path, verbose=False)

    detections = []
    for result in results:
        for box in result.boxes:
            class_id = int(box.cls[0])
            class_name = result.names[class_id]

            if class_name in FOOD_CALORIES:
                item, calories = FOOD_CALORIES[class_name]
                detections.append({"item": item, "calories": calories})

    if detections:
        return {"detections": detections}
    return None


if __name__ == "__main__":
    test_image_path = "test_image.jpg"
    try:
        from PIL import Image
        import numpy as np

        try:
            img = Image.open(test_image_path)
        except FileNotFoundError:
            print(f"Creating dummy test image: {test_image_path}")
            dummy_array = np.zeros((416, 416, 3), dtype=np.uint8)
            img = Image.fromarray(dummy_array)
            img.save(test_image_path)

        analysis_result = analyze_image(test_image_path)

        if analysis_result:
            print(f"Analysis Result: {analysis_result}")
        else:
            print("No food detected in the image.")

    except Exception as e:
        print(f"An error occurred: {e}")
