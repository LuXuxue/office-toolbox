"""FastAPI application for D-FINE layout analysis."""

from __future__ import annotations

import asyncio
from contextlib import asynccontextmanager
import time
from typing import Callable, List, Optional, TYPE_CHECKING

from fastapi import FastAPI, File, Form, Header, HTTPException, Response, UploadFile

from .confidence import format_predictions, parse_confidence_policy
from .config import Settings

if TYPE_CHECKING:
    from .runtime import ModelRuntime


def _server_timing(values: dict[str, float]) -> str:
    return ", ".join(f"{name};dur={duration:.6f}" for name, duration in values.items())


def create_app(
    runtime_factory: Optional[Callable[[Settings], "ModelRuntime"]] = None,
) -> FastAPI:
    @asynccontextmanager
    async def lifespan(app: FastAPI):
        settings = Settings.from_env()
        app.state.settings = settings
        factory = runtime_factory
        if factory is None:
            from .runtime import ModelRuntime

            factory = ModelRuntime
        app.state.runtime = factory(settings)
        yield

    app = FastAPI(
        title="D-FINE LA API",
        description="Inference API for the D-FINE Stage-2 B2 layout model",
        version="0.1.0",
        lifespan=lifespan,
        docs_url=None,
        redoc_url=None,
        openapi_url=None,
    )

    @app.post("/inference")
    async def inference(
        response: Response,
        images: List[UploadFile] = File(...),
        confidence: Optional[float] = Form(default=None),
        class_confidences: Optional[str] = Form(default=None),
        x_dfine_parallel_preprocess: Optional[str] = Header(default=None),
    ):
        runtime = app.state.runtime
        try:
            global_confidence, confidence_overrides, response_params = (
                parse_confidence_policy(
                    confidence,
                    class_confidences,
                    runtime.settings.num_classes,
                )
            )
        except ValueError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc
        request_started = time.perf_counter()
        payloads = [await image.read() for image in images]
        request_read_ms = (time.perf_counter() - request_started) * 1000
        parallel = (
            runtime.settings.parallel_preprocess
            if x_dfine_parallel_preprocess is None
            else x_dfine_parallel_preprocess == "1"
        )

        if parallel:
            loop = asyncio.get_running_loop()
            ready, decode_ms, preprocess_ms = await loop.run_in_executor(
                None, lambda: runtime.prepare(payloads)
            )
            outputs, inference_timings = await loop.run_in_executor(
                None, lambda: runtime.infer(ready)
            )
        else:
            ready, decode_ms, preprocess_ms = runtime.prepare(payloads)
            outputs, inference_timings = runtime.infer(ready)

        response.headers["Server-Timing"] = _server_timing(
            {
                "request_read": request_read_ms,
                "decode": decode_ms,
                "preprocess": preprocess_ms,
                **inference_timings,
            }
        )
        return format_predictions(
            outputs,
            global_confidence,
            confidence_overrides,
            response_params,
        )

    @app.get("/health")
    async def health():
        return app.state.runtime.health()

    @app.get("/model/info")
    async def model_info():
        return app.state.runtime.info()

    return app


app = create_app()
