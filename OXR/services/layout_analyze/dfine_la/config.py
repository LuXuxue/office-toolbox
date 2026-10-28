"""Fixed configuration and checkpoint validation for the LA model."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field, replace
import hashlib
import json
import os
from pathlib import Path
import re


# Capabilities of this implementation, not the identity of a specific checkpoint.
SUPPORTED_METADATA = {
    "schema_version": 1,
    "model_type": "dfine",
    "architecture": "stage2_b2",
    "num_classes": 12,
    "image_size": 960,
    "color_order": "rgb",
    "resize_backend": "opencv_lanczos4",
    "postprocess": "model_letterbox",
    "provides_reading_order": False,
    "id2label": {
        str(i): label for i, label in enumerate([
            "Caption", "Footnote", "Formula", "Text", "Page-footer", "Page-header",
            "Picture", "Title", "Table", "Code Block", "Stamp", "Chemical Structure",
        ])
    },
}


def _enabled(name: str, default: str = "1") -> bool:
    return os.getenv(name, default).strip().lower() in {"1", "true", "yes", "on"}


def _device(name: str = "DFINE_DEVICE", default: str = "cuda") -> str:
    value = os.getenv(name, default).strip().lower()
    if value not in {"auto", "cpu", "cuda"}:
        raise RuntimeError(
            f"{name} must be one of auto, cpu, or cuda; got {value!r}"
        )
    return value


@dataclass(frozen=True)
class Settings:
    model_path: Path
    metadata_path: Path | None = None
    model_version: str = "unknown"
    checkpoint_size_bytes: int | None = None
    checkpoint_sha256: str | None = None
    id2label: dict[str, str] = field(default_factory=dict)
    architecture: str = "stage2_b2"
    num_classes: int = 12
    image_size: int = 960
    color_order: str = "rgb"
    resize_backend: str = "opencv_lanczos4"
    postprocess: str = "model_letterbox"
    device: str = "cuda"
    inference_precision: str = "fp16_autocast"
    cuda_graph_enabled: bool = True
    parallel_preprocess: bool = True

    @classmethod
    def from_env(cls) -> "Settings":
        return cls(
            model_path=Path(os.getenv("MODEL_PATH", "/models/layout_analysis/last.pth")),
            metadata_path=(
                Path(os.environ["MODEL_METADATA_PATH"])
                if os.getenv("MODEL_METADATA_PATH") else None
            ),
            device=_device(),
            cuda_graph_enabled=_enabled("DFINE_CUDA_GRAPH"),
            parallel_preprocess=_enabled("DFINE_PARALLEL_PREPROCESS"),
        )

    def load_metadata(self) -> "Settings":
        """Load model-owned identity and reject unsupported inference semantics."""
        path = self.metadata_path or self.model_path.with_name("metadata.json")
        try:
            metadata = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            raise RuntimeError(f"unable to read model metadata: {path}") from exc
        if not isinstance(metadata, dict):
            raise RuntimeError("model metadata must be a JSON object")
        for name, supported in SUPPORTED_METADATA.items():
            value = metadata.get(name)
            if type(value) is not type(supported) or value != supported:
                raise RuntimeError(f"unsupported model metadata field: {name}")
        version = metadata.get("model_version")
        if not isinstance(version, str) or not re.fullmatch(r"\d+\.\d+\.\d+", version):
            raise RuntimeError("model metadata requires a public model_version (major.minor.patch)")
        # Do not resolve an arbitrary path from metadata; MODEL_PATH remains explicit.
        if metadata.get("checkpoint") != self.model_path.name:
            raise RuntimeError("model metadata checkpoint does not match MODEL_PATH filename")
        size = metadata.get("checkpoint_size_bytes")
        sha256 = metadata.get("checkpoint_sha256")
        if type(size) is not int or size <= 0:
            raise RuntimeError("model metadata requires positive checkpoint_size_bytes")
        if not isinstance(sha256, str) or not re.fullmatch(r"[0-9a-f]{64}", sha256):
            raise RuntimeError("model metadata requires a lowercase checkpoint_sha256")
        return replace(
            self, metadata_path=path, model_version=version,
            checkpoint_size_bytes=size, checkpoint_sha256=sha256,
            architecture=metadata["architecture"], num_classes=metadata["num_classes"],
            image_size=metadata["image_size"], color_order=metadata["color_order"],
            resize_backend=metadata["resize_backend"], postprocess=metadata["postprocess"],
            id2label=metadata["id2label"],
        )

    def validate_checkpoint(self) -> "Settings":
        if not self.model_path.is_file():
            raise RuntimeError(
                f"checkpoint not found: {self.model_path}; mount the model at this path"
            )
        declared = self.load_metadata()
        actual_size = self.model_path.stat().st_size
        if actual_size != declared.checkpoint_size_bytes:
            raise RuntimeError(
                f"checkpoint size mismatch: expected {declared.checkpoint_size_bytes}, got {actual_size}"
            )
        digest = hashlib.sha256()
        with self.model_path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(8 * 1024 * 1024), b""):
                digest.update(chunk)
        actual_sha256 = digest.hexdigest()
        if actual_sha256 != declared.checkpoint_sha256:
            raise RuntimeError(
                f"checkpoint SHA-256 mismatch: expected {declared.checkpoint_sha256}, got {actual_sha256}"
            )
        return declared

    def public_dict(self) -> dict:
        data = asdict(self)
        # Local mount locations are operational details, not public model metadata.
        data.pop("model_path")
        data.pop("metadata_path")
        return data
