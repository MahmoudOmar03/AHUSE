"""
Example one-off image call to the deployed YOLO API.
Configure YOLO_API_URL and YOLO_API_KEY in the environment (never commit secrets).
"""

import os
import sys

import requests


def main() -> None:
    url = os.environ.get("YOLO_API_URL", "").strip()
    key = os.environ.get("YOLO_API_KEY", "").strip()
    if not url or not key:
        print("Set YOLO_API_URL and YOLO_API_KEY.", file=sys.stderr)
        sys.exit(1)

    image_path = sys.argv[1] if len(sys.argv) > 1 else None
    if not image_path:
        print("Usage: python yolo_api_images.py <path-to-image>", file=sys.stderr)
        sys.exit(1)

    headers = {"Authorization": f"Bearer {key}"}
    conf = float(os.environ.get("YOLO_CONFIDENCE", "0.25"))
    iou = float(os.environ.get("YOLO_IOU", "0.7"))
    imgsz = int(os.environ.get("YOLO_IMGSZ", "640"))
    data = {"conf": conf, "iou": iou, "imgsz": imgsz}

    with open(image_path, "rb") as f:
        response = requests.post(
            url,
            headers=headers,
            data=data,
            files={"file": f},
            timeout=120,
        )
    response.raise_for_status()
    print(response.json())


if __name__ == "__main__":
    main()
