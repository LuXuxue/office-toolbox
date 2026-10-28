"""Native CUDA Graph replay for fixed-shape Batch1 inference."""

from __future__ import annotations

import torch


def _signature(images: torch.Tensor, sizes: torch.Tensor) -> tuple:
    return tuple(images.shape), images.dtype, tuple(sizes.shape), sizes.dtype


class _CapturedForward:
    def __init__(self, model, images, sizes, warmup_iterations: int) -> None:
        self.static_images = torch.empty_like(images)
        self.static_sizes = torch.empty_like(sizes)
        self.static_images.copy_(images)
        self.static_sizes.copy_(sizes)

        warmup_stream = torch.cuda.Stream(device=images.device)
        warmup_stream.wait_stream(torch.cuda.current_stream(images.device))
        with torch.cuda.stream(warmup_stream), torch.no_grad():
            for _ in range(warmup_iterations):
                model(self.static_images, self.static_sizes)
        torch.cuda.current_stream(images.device).wait_stream(warmup_stream)

        self.graph = torch.cuda.CUDAGraph()
        with torch.cuda.graph(self.graph), torch.no_grad():
            self.static_outputs = model(self.static_images, self.static_sizes)

    def __call__(self, images, sizes):
        self.static_images.copy_(images)
        self.static_sizes.copy_(sizes)
        self.graph.replay()
        return self.static_outputs


class CUDAGraphModel:
    def __init__(self, model, warmup_iterations: int = 3) -> None:
        self.model = model
        self.device = model.device
        self.warmup_iterations = warmup_iterations
        self._captures: dict[tuple, _CapturedForward] = {}

    def __call__(self, images: torch.Tensor, sizes: torch.Tensor):
        if images.shape[0] != 1:
            with torch.no_grad():
                return self.model(images, sizes)
        signature = _signature(images, sizes)
        capture = self._captures.get(signature)
        if capture is None:
            capture = _CapturedForward(
                self.model,
                images,
                sizes,
                self.warmup_iterations,
            )
            self._captures[signature] = capture
        return capture(images, sizes)

    def warmup(self, image_size: int) -> None:
        images = torch.zeros(
            (1, 3, image_size, image_size),
            dtype=torch.float32,
            device=self.device,
        )
        sizes = torch.full(
            (1, 2),
            image_size,
            dtype=torch.int64,
            device=self.device,
        )
        self(images, sizes)
        torch.cuda.synchronize(self.device)

    @property
    def ready(self) -> bool:
        return bool(self._captures)
