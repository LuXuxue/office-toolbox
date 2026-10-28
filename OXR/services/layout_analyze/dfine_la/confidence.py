"""Request-level confidence filtering for layout predictions."""

from __future__ import annotations

import json
import math
from numbers import Real

import numpy as np


def _validate_confidence(value: object, field: str) -> float:
    if isinstance(value, bool) or not isinstance(value, Real):
        raise ValueError(f"{field} must be a number between 0 and 1")
    confidence = float(value)
    if not math.isfinite(confidence) or not 0 <= confidence <= 1:
        raise ValueError(f"{field} must be a finite number between 0 and 1")
    return confidence


def parse_confidence_policy(
    confidence: float | None,
    class_confidences: str | None,
    num_classes: int,
) -> tuple[float, dict[int, float], dict[str, object]]:
    """Parse the global threshold and optional JSON class overrides."""
    global_confidence = (
        0.0
        if confidence is None
        else _validate_confidence(confidence, "confidence")
    )
    overrides: dict[int, float] = {}
    raw_overrides = class_confidences.strip() if class_confidences else ""
    if raw_overrides:
        try:
            parsed = json.loads(raw_overrides)
        except json.JSONDecodeError as exc:
            raise ValueError("class_confidences must be a JSON object") from exc
        if not isinstance(parsed, dict):
            raise ValueError("class_confidences must be a JSON object")
        for raw_label, raw_confidence in parsed.items():
            try:
                label = int(raw_label)
            except (TypeError, ValueError) as exc:
                raise ValueError(
                    f"class_confidences key {raw_label!r} must be a class id"
                ) from exc
            if str(label) != str(raw_label):
                raise ValueError(
                    f"class_confidences key {raw_label!r} must be a class id"
                )
            if not 0 <= label < num_classes:
                raise ValueError(
                    f"class_confidences class id {label} is outside [0, {num_classes - 1}]"
                )
            overrides[label] = _validate_confidence(
                raw_confidence,
                f"class_confidences[{label}]",
            )

    if confidence is None and not raw_overrides:
        return global_confidence, overrides, {}
    return global_confidence, overrides, {
        "confidence": global_confidence,
        "class_confidences": {
            str(label): threshold for label, threshold in sorted(overrides.items())
        },
    }


def format_predictions(
    outputs: tuple[np.ndarray, np.ndarray, np.ndarray],
    global_confidence: float = 0.0,
    class_confidences: dict[int, float] | None = None,
    params: dict[str, object] | None = None,
) -> dict:
    """Filter aligned labels, boxes, and scores without changing their order."""
    labels, boxes, scores = outputs
    overrides = class_confidences or {}
    if global_confidence == 0.0 and not overrides:
        filtered_labels = labels.tolist()
        filtered_boxes = boxes.tolist()
        filtered_scores = scores.tolist()
    else:
        filtered_labels = []
        filtered_boxes = []
        filtered_scores = []
        for image_labels, image_boxes, image_scores in zip(labels, boxes, scores):
            thresholds = np.full(image_scores.shape, global_confidence, dtype=np.float64)
            for label, threshold in overrides.items():
                thresholds[image_labels == label] = threshold
            keep = image_scores >= thresholds
            filtered_labels.append(image_labels[keep].tolist())
            filtered_boxes.append(image_boxes[keep].tolist())
            filtered_scores.append(image_scores[keep].tolist())

    return {
        "status": "success",
        "params": params or {},
        "data": {
            "labels": filtered_labels,
            "boxes": filtered_boxes,
            "scores": filtered_scores,
        },
    }
