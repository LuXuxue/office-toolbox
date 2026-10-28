"""Single-device model runtime with optional parallel preprocessing."""

from __future__ import annotations

import queue
import time

import torch

from .config import Settings
from .cuda_graph import CUDAGraphModel
from .model import load_model
from .preprocess import decode_images, preprocess_images


class ModelRuntime:
    def __init__(self, settings: Settings) -> None:
        settings = settings.validate_checkpoint()
        if settings.device == "cuda" and not torch.cuda.is_available():
            raise RuntimeError("an NVIDIA GPU with CUDA support is required")

        self.settings = settings
        device_name = (
            "cuda:0"
            if settings.device == "cuda"
            or (settings.device == "auto" and torch.cuda.is_available())
            else "cpu"
        )
        self.device = torch.device(device_name)
        model = load_model(
            settings.model_path,
            self.device,
            num_classes=settings.num_classes,
            image_size=settings.image_size,
        )
        self.cuda_graph_enabled = settings.cuda_graph_enabled and self.device.type == "cuda"
        if self.cuda_graph_enabled:
            model = CUDAGraphModel(model)
            model.warmup(settings.image_size)
        self._engines: queue.Queue = queue.Queue(maxsize=1)
        self._engines.put(model)
        self.cuda_graph_ready = bool(getattr(model, "ready", False))
        if self.cuda_graph_enabled and not self.cuda_graph_ready:
            raise RuntimeError("CUDA Graph warmup did not reach readiness")

    def prepare(self, payloads: list[bytes]):
        decode_started = time.perf_counter()
        images = decode_images(payloads)
        decode_ms = (time.perf_counter() - decode_started) * 1000

        preprocess_started = time.perf_counter()
        ready = preprocess_images(images, self.settings.image_size)
        preprocess_ms = (time.perf_counter() - preprocess_started) * 1000
        return ready, decode_ms, preprocess_ms

    def infer(self, ready):
        queue_started = time.perf_counter()
        model = self._engines.get(block=True)
        queue_ms = (time.perf_counter() - queue_started) * 1000
        try:
            images = ready[0].to(self.device)
            sizes = ready[1].to(self.device)
            inference_started = time.perf_counter()
            with torch.no_grad():
                labels, boxes, scores = model(images, sizes)
            outputs = tuple(
                value.cpu().detach().numpy() for value in (labels, boxes, scores)
            )
            # The CPU copy synchronizes CUDA. Include it so Server-Timing
            # reports completed inference instead of asynchronous launch time.
            inference_ms = (time.perf_counter() - inference_started) * 1000
            return outputs, {"queue": queue_ms, "inference": inference_ms}
        finally:
            self._engines.put(model)

    def health(self) -> dict:
        return {
            "status": "healthy",
            "model_version": self.settings.model_version,
            "num_classes": self.settings.num_classes,
            "img_size": self.settings.image_size,
            "preprocess_type": "resize_pad",
            "architecture": self.settings.architecture,
            "color_order": self.settings.color_order,
            "resize_backend": self.settings.resize_backend,
            "postprocess_mode": self.settings.postprocess,
            "device": str(self.device),
            "inference_precision": (
                "fp16_autocast" if self.device.type == "cuda" else "fp32"
            ),
            "cuda_graph_enabled": self.cuda_graph_enabled,
            "cuda_graph_ready": self.cuda_graph_ready,
            "cuda_graph_batch_sizes": [1],
        }

    def info(self) -> dict:
        model_config = self.settings.public_dict()
        model_config.update(
            device=str(self.device),
            inference_precision=(
                "fp16_autocast" if self.device.type == "cuda" else "fp32"
            ),
            cuda_graph_enabled=self.cuda_graph_enabled,
        )
        return {
            "model_config": model_config,
            "device_count": 1,
            "devices": [str(self.device)],
            "endpoint": "/inference",
        }
