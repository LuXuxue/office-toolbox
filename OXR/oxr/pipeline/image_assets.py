"""Request-scoped image output for the online demo; ordinary API output is unchanged."""

from contextvars import ContextVar
from dataclasses import dataclass
from pathlib import Path
from uuid import uuid4

import cv2
import numpy as np


@dataclass(frozen=True)
class ImageAssets:
    directory: Path
    url_prefix: str

    def save(self, image: np.ndarray) -> str:
        self.directory.mkdir(parents=True, exist_ok=True)
        name = f"block-{uuid4().hex}.png"
        ok, encoded = cv2.imencode(".png", image)
        if not ok:
            raise ValueError("Unable to encode an output image.")
        (self.directory / name).write_bytes(encoded.tobytes())
        return self.url_prefix.rstrip("/") + "/" + name


image_assets: ContextVar[ImageAssets | None] = ContextVar("oxr_image_assets", default=None)
