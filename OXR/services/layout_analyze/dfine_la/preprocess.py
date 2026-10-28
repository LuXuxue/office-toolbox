"""OpenCV preprocessing equivalent to the official Stage-2 evaluator."""

from __future__ import annotations

import cv2
import numpy as np
import torch


def decode_images(payloads: list[bytes]) -> list[np.ndarray]:
    images = []
    for payload in payloads:
        image = cv2.imdecode(np.frombuffer(payload, dtype=np.uint8), cv2.IMREAD_COLOR)
        if image is None:
            raise ValueError("OpenCV failed to decode an uploaded image")
        images.append(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))
    return images


def resize_and_pad(
    image: np.ndarray,
    target_size: int = 960,
    fill_value: int = 114,
) -> np.ndarray:
    original_height, original_width = image.shape[:2]
    ratio = target_size / max(original_width, original_height)
    new_width = int(original_width * ratio)
    new_height = int(original_height * ratio)
    pad_left = (target_size - new_width) // 2
    pad_top = (target_size - new_height) // 2
    resized = cv2.resize(
        image,
        (new_width, new_height),
        interpolation=cv2.INTER_LANCZOS4,
    )
    canvas = np.full(
        (target_size, target_size, resized.shape[2]),
        fill_value,
        dtype=image.dtype,
    )
    canvas[pad_top : pad_top + new_height, pad_left : pad_left + new_width] = resized
    return canvas


def uint8_hwc_to_tensor(image: np.ndarray) -> torch.Tensor:
    if image.dtype != np.uint8 or image.ndim != 3:
        raise ValueError("expected an HWC uint8 image")
    chw = np.ascontiguousarray(image.transpose((2, 0, 1)))
    return torch.from_numpy(chw).to(dtype=torch.get_default_dtype()).div(255)


def preprocess_images(images: list[np.ndarray], image_size: int = 960):
    tensors = []
    original_sizes = []
    for image in images:
        height, width = image.shape[:2]
        tensors.append(uint8_hwc_to_tensor(resize_and_pad(image, image_size)))
        original_sizes.append([width, height])
    batch = tensors[0].unsqueeze(0) if len(tensors) == 1 else torch.stack(tensors)
    return batch, torch.tensor(original_sizes)
