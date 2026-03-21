"""
Sample frames from a video and run each through the YOLO predict API.
Uses the same env config as yolo_api_client (YOLO_API_URL, YOLO_API_KEY).
"""

from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Optional

import cv2
import requests

from yolo_api_client import get_yolo_config, parse_prediction_payload


def detect_objects_in_video(
    video_path: str,
    sample_every_sec: int = 1,
    conf: Optional[float] = None,
    iou: Optional[float] = None,
    imgsz: Optional[int] = None,
    timeout: int = 120,
) -> Dict[str, Any]:
    url, headers, conf_d, iou_d, imgsz_d = get_yolo_config()
    if conf is not None:
        conf_d = conf
    if iou is not None:
        iou_d = iou
    if imgsz is not None:
        imgsz_d = imgsz

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        raise ValueError(f"Could not open video: {video_path}")

    fps = cap.get(cv2.CAP_PROP_FPS)
    if not fps or fps <= 0:
        fps = 25.0

    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    duration_sec = total_frames / fps if total_frames > 0 else 0

    frame_step = max(1, int(round(fps * sample_every_sec)))

    per_second_results: List[Dict[str, Any]] = []
    all_detected_names: List[str] = []
    name_counter: Counter = Counter()

    frame_idx = 0
    sampled_index = 0

    data = {"conf": conf_d, "iou": iou_d, "imgsz": imgsz_d}

    while True:
        ok, frame = cap.read()
        if not ok:
            break

        if frame_idx % frame_step == 0:
            second_mark = int(frame_idx / fps)

            success, encoded = cv2.imencode(".jpg", frame)
            if not success:
                frame_idx += 1
                continue

            files = {"file": ("frame.jpg", encoded.tobytes(), "image/jpeg")}
            response = requests.post(
                url,
                headers=headers,
                data=data,
                files=files,
                timeout=timeout,
            )
            response.raise_for_status()
            payload = response.json()
            detections_raw = parse_prediction_payload(payload)

            frame_objects = []
            for det in detections_raw:
                name = det.get("class_name", "")
                confidence = det.get("confidence", 0.0)
                box_list = det.get("bbox") or []
                box_dict = (
                    {
                        "x1": box_list[0],
                        "y1": box_list[1],
                        "x2": box_list[2],
                        "y2": box_list[3],
                    }
                    if len(box_list) >= 4
                    else {}
                )
                class_id = None
                frame_objects.append(
                    {
                        "class_id": class_id,
                        "name": name,
                        "confidence": confidence,
                        "box": box_dict,
                    },
                )
                if name:
                    all_detected_names.append(name)
                    name_counter[name] += 1

            per_second_results.append(
                {
                    "sample_index": sampled_index,
                    "second": second_mark,
                    "frame_index": frame_idx,
                    "detections": frame_objects,
                },
            )
            sampled_index += 1

        frame_idx += 1

    cap.release()

    return {
        "video_path": video_path,
        "fps": fps,
        "duration_sec": round(duration_sec, 2),
        "sample_every_sec": sample_every_sec,
        "total_samples": len(per_second_results),
        "unique_detected_objects": sorted(set(all_detected_names)),
        "object_counts": dict(name_counter),
        "per_second_results": per_second_results,
    }


if __name__ == "__main__":
    import sys

    vp = sys.argv[1] if len(sys.argv) > 1 else "uploads/my_video.mp4"
    result = detect_objects_in_video(video_path=vp, sample_every_sec=1)
    print("\n=== UNIQUE DETECTED OBJECTS ===")
    print(result["unique_detected_objects"])
    print("\n=== OBJECT COUNTS ===")
    print(result["object_counts"])
