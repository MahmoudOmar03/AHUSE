import json
import shutil
from dataclasses import dataclass
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from typing import Any, Dict, List, Optional

from LLM_VLM import generate_hse_report

try:
    from ultralytics import YOLO
except Exception as exc:  # pragma: no cover
    YOLO = None  # type: ignore[assignment]
    _yolo_import_error = exc
else:
    _yolo_import_error = None

from yolo_api_client import predict_image_file

# Keywords that indicate a relevant construction / PPE detection.
RELEVANT_KEYWORDS = {
    "helmet",
    "hardhat",
    "hard hat",
    "vest",
    "ppe",
    "glove",
    "harness",
    "worker",
    "person",
    "construction",
    "machinery",
    "crane",
    "scaffold",
}

VIDEO_YOLO_JSON = "video_yolo_summary.json"


@dataclass
class YoloDetection:
    class_name: str
    confidence: float
    bbox: List[float]


def _ensure_ultralytics():
    if YOLO is None:
        raise RuntimeError(
            "ultralytics is not installed. Install it with `pip install ultralytics`.",
        ) from _yolo_import_error


@lru_cache(maxsize=1)
def _load_yolo(model_path: str) -> Any:
    _ensure_ultralytics()
    model_path_resolved = Path(model_path)
    if not model_path_resolved.exists():
        raise FileNotFoundError(f"YOLO model weights not found: {model_path}")
    return YOLO(model_path_resolved)


def run_yolo_gate_local(
    image_path: str,
    model_path: str,
    confidence_threshold: float = 0.25,
) -> Dict[str, Any]:
    model = _load_yolo(model_path)
    results = model(image_path, conf=confidence_threshold)

    detections: List[YoloDetection] = []
    relevant = False

    for result in results:
        names = result.names
        boxes = getattr(result, "boxes", None)
        if boxes is None:
            continue
        for box in boxes:
            cls_idx = int(box.cls)
            class_name = names.get(cls_idx, str(cls_idx))
            conf = float(box.conf)
            if conf < confidence_threshold:
                continue
            bbox = box.xyxy[0].tolist()
            detections.append(
                YoloDetection(
                    class_name=class_name,
                    confidence=round(conf, 3),
                    bbox=[round(float(coord), 2) for coord in bbox],
                ),
            )
            if _is_relevant_detection(class_name):
                relevant = True

    return {
        "relevant": relevant,
        "detections": [det.__dict__ for det in detections],
    }


def run_yolo_gate_api(
    image_path: str,
    confidence_threshold: float = 0.25,
) -> Dict[str, Any]:
    raw = predict_image_file(image_path)
    detections: List[Dict[str, Any]] = []
    relevant = False
    for det in raw:
        conf = float(det.get("confidence", 0))
        if conf < confidence_threshold:
            continue
        class_name = det.get("class_name", "")
        bbox = det.get("bbox") or []
        detections.append(
            {
                "class_name": class_name,
                "confidence": round(conf, 4),
                "bbox": bbox,
            },
        )
        if _is_relevant_detection(class_name):
            relevant = True
    return {"relevant": relevant, "detections": detections}


def merge_yolo_gates(gates: List[Dict[str, Any]]) -> Dict[str, Any]:
    merged: List[Dict[str, Any]] = []
    relevant = False
    for g in gates:
        if g.get("relevant"):
            relevant = True
        merged.extend(g.get("detections") or [])
    return {"relevant": relevant, "detections": merged}


def run_yolo_for_images(
    config: Dict[str, Any],
    image_paths: List[str],
) -> Dict[str, Any]:
    confidence_threshold = float(config.get("YOLO_CONFIDENCE", 0.25))
    use_local = str(config.get("YOLO_USE_LOCAL", "")).lower() in ("1", "true", "yes")
    yolo_weights = config.get("YOLO_MODEL_PATH", "")

    gates: List[Dict[str, Any]] = []
    for p in image_paths:
        if use_local:
            gates.append(
                run_yolo_gate_local(
                    image_path=p,
                    model_path=str(yolo_weights),
                    confidence_threshold=confidence_threshold,
                ),
            )
        else:
            gates.append(
                run_yolo_gate_api(
                    image_path=p,
                    confidence_threshold=confidence_threshold,
                ),
            )
    return merge_yolo_gates(gates)


def _is_relevant_detection(class_name: str) -> bool:
    lower = class_name.lower()
    return any(keyword in lower for keyword in RELEVANT_KEYWORDS)


def _video_names_relevant(unique_names: List[str]) -> bool:
    return any(_is_relevant_detection(n) for n in unique_names)


def process_hse_request(
    config: Dict[str, Any],
    image_path: str,
    project_name: str,
    site_location: str,
    inspection_by: str,
    verified_by: str,
    max_records: Optional[int] = None,
    image_paths_for_yolo: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Full pipeline for image-based inspection. Primary image_path is used for the VLM report;
    image_paths_for_yolo defaults to [image_path] for merged YOLO relevance and detections.
    """
    outputs_base = Path(config["OUTPUT_FOLDER"])
    max_records = max_records or int(config.get("MAX_RECORDS", 5))

    yolo_paths = image_paths_for_yolo if image_paths_for_yolo else [image_path]

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    request_folder = outputs_base / timestamp
    request_folder.mkdir(parents=True, exist_ok=True)

    for p in yolo_paths:
        shutil.copy2(p, request_folder / Path(p).name)

    try:
        gate = run_yolo_for_images(config, yolo_paths)
    except Exception as exc:
        return {
            "status": "error",
            "message": f"YOLO gate failed: {exc}",
            "output_folder": request_folder.name,
        }

    if not gate["relevant"]:
        return {
            "status": "not_relevant",
            "message": "No construction or PPE-related activity detected. Report skipped.",
            "detections": gate["detections"],
            "output_folder": request_folder.name,
        }

    try:
        report = generate_hse_report(
            image_path=image_path,
            project_name=project_name,
            site_location=site_location,
            inspection_by=inspection_by,
            verified_by=verified_by,
            out_dir=str(request_folder),
            max_records=max_records,
        )
    except Exception as exc:
        return {
            "status": "error",
            "message": f"Failed to generate HSE report: {exc}",
            "detections": gate["detections"],
            "output_folder": request_folder.name,
        }

    return {
        "status": "ok",
        "message": "Report generated successfully.",
        "detections": gate["detections"],
        "output_folder": request_folder.name,
        "hse": report["hse"],
        "json_path": report["json_path"],
        "docx_path": report["docx_path"],
        "raw_text_path": report["raw_text_path"],
        "raw_model_output": report["raw_model_output"],
        "generated_at": report["generated_at"],
    }


def process_video_only_request(
    config: Dict[str, Any],
    video_path: str,
    sample_every_sec: int = 1,
) -> Dict[str, Any]:
    """Create output folder, copy video, run API sampling, write video_yolo_summary.json."""
    from yolo_api_video import detect_objects_in_video

    outputs_base = Path(config["OUTPUT_FOLDER"])
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    request_folder = outputs_base / timestamp
    request_folder.mkdir(parents=True, exist_ok=True)

    vpath = Path(video_path)
    dest_video = request_folder / vpath.name
    shutil.copy2(video_path, dest_video)

    try:
        summary = detect_objects_in_video(
            str(dest_video),
            sample_every_sec=sample_every_sec,
        )
    except Exception as exc:
        return {
            "status": "error",
            "message": f"Video inference failed: {exc}",
            "output_folder": request_folder.name,
        }

    summary_path = request_folder / VIDEO_YOLO_JSON
    summary_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")

    relevant = _video_names_relevant(summary.get("unique_detected_objects") or [])

    return {
        "status": "ok",
        "message": "Video analysis completed.",
        "output_folder": request_folder.name,
        "video_summary_path": str(summary_path),
        "video_relevant": relevant,
        "video_summary": summary,
    }


def append_video_summary_to_folder(
    output_folder_name: str,
    config: Dict[str, Any],
    video_path: str,
    sample_every_sec: int = 1,
) -> Optional[str]:
    """Run video inference and write video_yolo_summary.json into an existing output folder."""
    from yolo_api_video import detect_objects_in_video

    outputs_base = Path(config["OUTPUT_FOLDER"])
    request_folder = outputs_base / output_folder_name
    if not request_folder.is_dir():
        return None

    vpath = Path(video_path)
    dest_video = request_folder / vpath.name
    shutil.copy2(video_path, dest_video)

    summary = detect_objects_in_video(
        str(dest_video),
        sample_every_sec=sample_every_sec,
    )
    summary_path = request_folder / VIDEO_YOLO_JSON
    summary_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    return str(summary_path)
