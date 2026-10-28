import asyncio
from typing import Any, Dict, List, Optional, Sequence

import numpy as np

from oxr.config.settings import LayoutAnalyzeModelConfig, settings
from oxr.model.layout_client import LayoutAnalyzeClient, layout_client
from oxr.pipeline.concurrency import gather_cancel_on_error
from oxr.pipeline.errors import PipelineError, format_error_message, get_error_status_code
from oxr.pipeline.models import Block
from oxr.pipeline.postprocess.layout_postprocess import parse_layout_batch


async def process_layout_batch(
    encoded_images: Sequence[bytes],
    image_sizes: Sequence[tuple[int, int]],
    sem: asyncio.Semaphore,
    batch_index: int,
    *,
    client: LayoutAnalyzeClient,
    config: LayoutAnalyzeModelConfig,
) -> list[list[Block]]:
    """Request one layout batch, retrying the entire batch when it fails."""
    for attempt in range(config.retry_times):
        try:
            # Only the upstream HTTP attempt is scarce.  Parsing and backoff do
            # not occupy a layout-service permit.
            async with sem:
                payload = await client.request_batch(encoded_images)
            return parse_layout_batch(
                payload,
                image_sizes,
                config.score_threshold,
                config.nms_iou_threshold,
            )
        except Exception as error:
            if attempt == config.retry_times - 1:
                message = format_error_message(error, f"Timeout after {config.timeout}s")
                if isinstance(error, TimeoutError) and not str(error).strip():
                    message = f"Timeout after {config.timeout}s"
                raise PipelineError(
                    f"Layout analysis failed for batch {batch_index} after "
                    f"{config.retry_times} attempts: {message}",
                    status_code=get_error_status_code(error, default=503),
                ) from error
            await asyncio.sleep(1)


async def analyze_layouts(
    images: List[np.ndarray],
    model_options: Optional[Dict[str, Any]] = None,
    *,
    client: Optional[LayoutAnalyzeClient] = None,
    config: Optional[LayoutAnalyzeModelConfig] = None,
) -> List[List[Block]]:
    """Analyze page layouts in concurrent, order-preserving service batches."""
    del model_options  # LayoutAnalyzeClient has a fixed structured API.
    if not images:
        return []

    resolved_config = config or (client.config if client is not None else settings.layout_analyze_model)
    owns_client = client is None and config is not None
    resolved_client = client or (
        LayoutAnalyzeClient(resolved_config) if owns_client else layout_client
    )
    encoded_images = resolved_client.encode_images(images)
    image_sizes = [(int(image.shape[1]), int(image.shape[0])) for image in images]
    batches = [
        (
            encoded_images[start : start + resolved_config.batch_size],
            image_sizes[start : start + resolved_config.batch_size],
        )
        for start in range(0, len(images), resolved_config.batch_size)
    ]
    sem = asyncio.Semaphore(resolved_config.max_concurrency)
    try:
        batch_results = await gather_cancel_on_error(
            *(
                process_layout_batch(
                    batch,
                    sizes,
                    sem,
                    batch_index,
                    client=resolved_client,
                    config=resolved_config,
                )
                for batch_index, (batch, sizes) in enumerate(batches)
            )
        )
        return [page for batch in batch_results for page in batch]
    finally:
        if owns_client:
            await resolved_client.aclose()
