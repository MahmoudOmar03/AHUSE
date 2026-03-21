"""
HTTP client for the deployed YOLO predict API (same contract as yolo_api_images.py).
Configuration via YOLO_API_URL and YOLO_API_KEY environment variables.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests


def get_yolo_config() -> Tuple[str, Dict[str, str], float, float, int]:
    url = os.environ.get("YOLO_API_URL", "").strip()
    key = os.environ.get("YOLO_API_KEY", "").strip()
    if not url or not key:
        raise RuntimeError(
            "YOLO_API_URL and YOLO_API_KEY must be set for cloud YOLO inference.",
        )
    headers = {"Authorization": f"Bearer {key}"}
    conf = float(os.environ.get("YOLO_CONFIDENCE", "0.25"))
    iou = float(os.environ.get("YOLO_IOU", "0.7"))
    imgsz = int(os.environ.get("YOLO_IMGSZ", "640"))
    return url, headers, conf, iou, imgsz


def _normalize_confidence(raw: Any) -> float:
    if raw is None:
        return 0.0
    try:
        v = float(raw)
    except (TypeError, ValueError):
        return 0.0
    if v > 1.0:
        v = v / 100.0
    return round(min(max(v, 0.0), 1.0), 4)


def _box_to_xyxy(box: Any) -> List[float]:
    if box is None:
        return []
    if isinstance(box, (list, tuple)) and len(box) >= 4:
        return [round(float(box[i]), 2) for i in range(4)]
    if isinstance(box, dict):
        for keys in (
            ("x1", "y1", "x2", "y2"),
            ("left", "top", "right", "bottom"),
            ("xmin", "ymin", "xmax", "ymax"),
        ):
            if all(k in box for k in keys):
                return [round(float(box[keys[i]]), 2) for i in range(4)]
    return []


def parse_prediction_payload(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Map API JSON to pipeline detection dicts: class_name, confidence, bbox."""
    out: List[Dict[str, Any]] = []
    images = payload.get("images") or []
    if not images:
        return out
    results = images[0].get("results") or []
    for det in results:
        name = det.get("name") or det.get("class_name") or ""
        conf = _normalize_confidence(det.get("confidence"))
        bbox = _box_to_xyxy(det.get("box") or det.get("bbox"))
        out.append(
            {
                "class_name": str(name),
                "confidence": conf,
                "bbox": bbox,
            },
        )
    return out


def predict_image_file(
    image_path: str,
    timeout: int = 120,
) -> List[Dict[str, Any]]:
    url, headers, conf, iou, imgsz = get_yolo_config()
    path = Path(image_path)
    data = {"conf": conf, "iou": iou, "imgsz": imgsz}
    with path.open("rb") as f:
        response = requests.post(
            url,
            headers=headers,
            data=data,
            files={"file": (path.name, f, "application/octet-stream")},
            timeout=timeout,
        )
    response.raise_for_status()
    return parse_prediction_payload(response.json())


def predict_image_bytes(
    filename: str,
    content: bytes,
    content_type: Optional[str] = None,
    timeout: int = 120,
) -> List[Dict[str, Any]]:
    url, headers, conf, iou, imgsz = get_yolo_config()
    data = {"conf": conf, "iou": iou, "imgsz": imgsz}
    mime = content_type or "image/jpeg"
    response = requests.post(
        url,
        headers=headers,
        data=data,
        files={"file": (filename, content, mime)},
        timeout=timeout,
    )
    response.raise_for_status()
    return parse_prediction_payload(response.json())
