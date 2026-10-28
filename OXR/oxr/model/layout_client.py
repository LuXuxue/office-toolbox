from typing import Optional, Sequence

import cv2
import httpx
import numpy as np

from oxr.config.settings import LayoutAnalyzeModelConfig, settings


class LayoutAnalyzeClient:
    """Persistent HTTP client for the D-FINE layout-analysis service."""

    def __init__(
        self,
        config: Optional[LayoutAnalyzeModelConfig] = None,
        *,
        http_client: Optional[httpx.AsyncClient] = None,
    ):
        self._config = config
        self._client = http_client

    @property
    def config(self) -> LayoutAnalyzeModelConfig:
        return self._config or settings.layout_analyze_model

    @property
    def client(self) -> httpx.AsyncClient:
        if self._client is None:
            if not self.config.url:
                raise RuntimeError("Layout analysis model URL is not configured.")
            self._client = httpx.AsyncClient(base_url=self.config.url, trust_env=False)
        return self._client

    def encode_images(self, images: Sequence[np.ndarray]) -> list[bytes]:
        encoded_images: list[bytes] = []
        for image in images:
            success, encoded = cv2.imencode(".png", image)
            if not success:
                raise ValueError("Unable to encode image as PNG.")
            encoded_images.append(encoded.tobytes())
        return encoded_images

    async def request_batch(self, encoded_images: Sequence[bytes]) -> object:
        files = [
            ("images", (f"page-{index}.png", image, "application/octet-stream"))
            for index, image in enumerate(encoded_images)
        ]
        base_url = str(self.config.url).rstrip("/")
        if base_url.endswith("/v1/layout/analyze"):
            base_url = base_url[: -len("/v1/layout/analyze")]
        endpoint = base_url if base_url.endswith("/inference") else f"{base_url}/inference"
        response = await self.client.post(endpoint, files=files, timeout=self.config.timeout)
        response.raise_for_status()
        return response.json()

    async def aclose(self) -> None:
        if self._client is not None:
            await self._client.aclose()
            self._client = None


layout_client = LayoutAnalyzeClient()
