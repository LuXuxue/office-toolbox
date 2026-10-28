import os
from typing import Any, Optional
from urllib.parse import urlparse

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator
from pydantic_settings import BaseSettings, SettingsConfigDict
import yaml


def _parse_env_bool(value: Optional[str], default: bool = False) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"true", "1", "yes", "on"}


def _validate_http_url(value: Optional[str]) -> Optional[str]:
    if value is None:
        return value
    parsed = urlparse(value)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("must be an absolute http:// or https:// URL")
    return value


class ServerConfig(BaseModel):
    model_config = ConfigDict(validate_assignment=True)

    host: str = "0.0.0.0"
    port: int = Field(default=8000, gt=0, le=65535)
    workers: int = Field(default=1, gt=0)


class LayoutAnalyzeModelConfig(BaseModel):
    model_config = ConfigDict(validate_assignment=True)

    url: Optional[str] = None
    batch_size: int = Field(default=8, gt=0)
    max_concurrency: int = Field(default=2, gt=0)
    timeout: int = Field(default=60, gt=0)
    retry_times: int = Field(default=2, gt=0)
    score_threshold: float = Field(default=0.5, ge=0.0, le=1.0)
    nms_iou_threshold: Optional[float] = Field(default=0.6, ge=0.0, le=1.0)
    provides_reading_order: bool = False

    _validate_url = field_validator("url")(_validate_http_url)


class OXRModelConfig(BaseModel):
    model_config = ConfigDict(validate_assignment=True)

    url: Optional[str] = None
    model_name: str = "oxr"
    max_tokens: int = Field(default=8192, gt=0)
    max_concurrency: int = Field(default=16, gt=0)
    timeout: int = Field(default=120, gt=0)
    retry_times: int = Field(default=2, gt=0)

    _validate_url = field_validator("url")(_validate_http_url)


class PipelineConfig(BaseModel):
    model_config = ConfigDict(validate_assignment=True, extra="forbid")

    image_output_dir: str = "outputs"
    full_overlap_dedupe: bool = False
    table_image_placeholders: bool = False
    table_image_min_score: float = Field(default=0.5, ge=0.0, le=1.0)


class Settings(BaseSettings):
    debug: bool = Field(default_factory=lambda: _parse_env_bool(os.getenv("DEBUG"), False))
    server: ServerConfig = Field(default_factory=ServerConfig)
    layout_analyze_model: LayoutAnalyzeModelConfig = Field(default_factory=LayoutAnalyzeModelConfig)
    oxr_model: OXRModelConfig = Field(default_factory=OXRModelConfig)
    pipeline: PipelineConfig = Field(default_factory=PipelineConfig)

    model_config = SettingsConfigDict(
        env_prefix="OXR__",
        env_nested_delimiter="__",
        extra="forbid",
    )

    @model_validator(mode="before")
    @classmethod
    def reject_legacy_model_configuration(cls, data: Any) -> Any:
        if isinstance(data, dict) and "model" in data:
            raise ValueError(
                "The legacy 'model' configuration is no longer supported; "
                "migrate it to 'layout_analyze_model' and 'oxr_model'."
            )

        legacy_environment = sorted(
            name
            for name in os.environ
            if name.upper().startswith("OXR_MODEL__") or name.upper().startswith("OXR__MODEL__")
        )
        if legacy_environment:
            raise ValueError(
                "legacy model configuration environment variables are not supported "
                f"({', '.join(legacy_environment)}); migrate to "
                "OXR__LAYOUT_ANALYZE_MODEL__* and OXR__OXR_MODEL__*."
            )
        return data

    def validate_for_server(self) -> None:
        missing = []
        if not self.layout_analyze_model.url:
            missing.append("layout_analyze_model.url")
        if not self.oxr_model.url:
            missing.append("oxr_model.url")
        if missing:
            raise ValueError(
                "Both model URLs are required to start the server: " + ", ".join(missing)
            )

    @classmethod
    def load_from_yaml(cls, yaml_path: str) -> "Settings":
        with open(yaml_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        return cls(**data)


settings = Settings()
