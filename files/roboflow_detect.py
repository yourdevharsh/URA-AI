from roboflow import Roboflow
import cv2
import os

# Initialize Roboflow model once
api_key = os.environ.get("ROBOFLOW_API_KEY", "He6P4fZ5RBJLwsdCxa4N")
try:
    rf = Roboflow(api_key=api_key)
    project = rf.workspace("ui-elements-t9wim").project("ui-elements-u0wsn")
    model = project.version("6").model  # Version 6
except Exception as e:
    print("❌ Roboflow initialization failed:", e)
    rf = None
    model = None

def detect_objects(image_path, confidence=0.50, save_annotated_path=None, filter_labels=None):
    """
    Detect UI elements in a given image using Roboflow online model.
    Only returns objects in `filter_labels` if provided.
    """
    prediction = model.predict(image_path, confidence=confidence).json()
    detections = []

    img = cv2.imread(image_path)
    if img is None:
        print("❌ Failed to read image:", image_path)
        return []

    for obj in prediction['predictions']:
        label = obj['class']
        if filter_labels and label not in filter_labels:
            continue

        x, y, w, h = obj['x'], obj['y'], obj['width'], obj['height']
        conf = obj['confidence']

        # Convert center coords to top-left / bottom-right
        x1 = int(x - w/2)
        y1 = int(y - h/2)
        x2 = int(x + w/2)
        y2 = int(y + h/2)

        detections.append({
            'label': label,
            'box': [x1, y1, x2, y2],
            'confidence': conf
        })

        if save_annotated_path:
            cv2.rectangle(img, (x1,y1), (x2,y2), (0,122,255), 2)
            cv2.putText(img, label, (x1, y1-5), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255,255,255), 1)

    if save_annotated_path:
        cv2.imwrite(save_annotated_path, img)

    return detections
