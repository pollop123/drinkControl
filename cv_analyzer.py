
from ultralytics import YOLO

# The model is now initialized inside the function to support lazy loading.
model = None

def analyze_image(image_path):
    """
    Analyzes an image using YOLOv8 to detect objects.
    The model is loaded on the first call to this function.

    Args:
        image_path (str): The path to the image file.

    Returns:
        dict: A dictionary with item and calorie information 
              if a 'bottle' is detected, otherwise None.
    """
    global model
    # Initialize the model only if it hasn't been initialized yet.
    if model is None:
        print("Initializing YOLO model for the first time...")
        model = YOLO('yolov8n.pt')
        print("YOLO model initialized.")

    # Run prediction on the image
    results = model.predict(image_path, verbose=False)

    # Check the results
    for result in results:
        # The names dictionary maps class IDs to class names
        class_names = result.names
        for box in result.boxes:
            class_id = int(box.cls[0])
            class_name = class_names[class_id]
            
            # For our MVP, we only care about 'bottle'
            if class_name == 'bottle':
                # Return a hardcoded calorie value for the MVP
                return {"item": "瓶裝飲料", "calories": 150}

    # If no bottle was found after checking all detections
    return None

if __name__ == '__main__':
    # This is for testing the module directly
    # You can replace 'test_image.jpg' with a path to an actual image
    # containing a bottle to test the functionality.
    test_image_path = 'test_image.jpg' 
    try:
        # Create a dummy image for testing if it doesn't exist
        from PIL import Image
        import numpy as np
        
        try:
            img = Image.open(test_image_path)
        except FileNotFoundError:
            print(f"Creating a dummy test image: {test_image_path}")
            dummy_array = np.zeros((416, 416, 3), dtype=np.uint8)
            img = Image.fromarray(dummy_array)
            img.save(test_image_path)

        analysis_result = analyze_image(test_image_path)

        if analysis_result:
            print(f"Analysis Result: {analysis_result}")
        else:
            print("No bottle detected in the image.")

    except ImportError:
        print("Please install Pillow and numpy to run the test part of this script.")
    except Exception as e:
        print(f"An error occurred: {e}")
